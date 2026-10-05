"""`bli2prism qc`: a QC report for one run, as a single self-contained HTML file.

Every check ends as PASS, WARN or FAIL with a plain-language reason. FAIL means a build would
be wrong or would stop; WARN means the numbers are usable but need a human look.

Thresholds (also printed in the report):
    CSV vs raw at the association time   WARN if |difference| > 0.005 nm
    equilibrium fit                      WARN if R² < 0.95, top plateau not reached (top conc < 10×EC50),
                                         EC50 outside the tested range, Top > 2× the highest response,
                                         or the LogEC50 interval is wider than 1 log unit
    equilibrium values                   WARN if a lower concentration gives >20% more response than the top one
    raw traces                           WARN if a trace ends association >10% below its peak (peak > 0.05 nm)
    kinetic fits                         see kinetics.py (X² < 3, R² > 0.9, determinate KD/kon/kdis;
                                         flagged = KD > 5-fold from the group median)
    KD estimates                         WARN if mean kdis/mean kon and the mean of per-row KDs
                                         differ by more than 2-fold
"""

from __future__ import annotations

import base64
import html
import io
import os
import statistics

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402

from . import eqfit
from . import kinetics as K
from . import refdata as R

CSV_RAW_WARN_NM = 0.005
KD_DISAGREE_FOLD = 2.0
HOOK_FACTOR = 1.2             # WARN if some concentration gives >20% more response than the top one
ASSOC_DECLINE = 0.10          # WARN if the end of association is >10% below its peak
ASSOC_MIN_SIGNAL_NM = 0.05    # ...and the peak is above the noise floor
LEVELS = {"FAIL": 0, "WARN": 1, "PASS": 2, "INFO": 3}


class Checks:
    def __init__(self):
        self.items = []

    def add(self, level, area, title, detail=""):
        self.items.append({"level": level, "area": area, "title": title, "detail": detail})

    def count(self, level):
        return sum(1 for c in self.items if c["level"] == level)


def _png(fig, dpi=110):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()


def _floats(col):
    return np.array([float(v) if v not in ("", None) else np.nan for v in col])


# --------------------------------------------------------------------------- the run
def run_qc(folder, setup_path, results_csv, n_conc, template=None, scratch=False, container=None):
    ck = Checks()
    res = {"checks": ck, "folder": folder, "setup": setup_path, "csv": results_csv, "n_conc": n_conc}
    rows = R.read_results_csv(results_csv)
    try:
        setup = R.read_setup(setup_path, {r["Loading Sample ID"] for r in rows})
    except ValueError as e:
        ck.add("FAIL", "Inputs", "Setup is not usable", str(e))
        return res
    res["setup"] = setup
    ligands, analytes, series = setup["ligands"], setup["analytes"], setup["series_M"]
    names = [a["csv_id"] for a in analytes]
    res["analytes"] = analytes
    ck.add("PASS", "Inputs", "Setup read",
           f"{len(ligands)} ligand(s) {ligands}; reference sensor {setup['reference']!r}; "
           f"{len(analytes)} analyte(s); {len(series)} dilution steps; normalization reference "
           f"{setup['reference_analyte']!r}")

    # --- results CSV against the setup
    extra = sorted({r["Sample ID"] for r in rows} - set(names))
    if extra:
        ck.add("FAIL", "Inputs", "CSV has analytes the setup does not list", str(extra))
    missing = sorted(set(names) - {r["Sample ID"] for r in rows})
    if missing:
        ck.add("FAIL", "Inputs", "Setup lists analytes absent from the CSV", str(missing))
    assoc = [k for k in rows[0] if k.startswith("X=") and k != "X=0"]
    if len(assoc) != 1:
        ck.add("FAIL", "Inputs", "Cannot identify the equilibrium column", f"found {assoc}")
        return res
    t_assoc = float(assoc[0][2:])
    res["t_assoc"] = t_assoc
    ck.add("PASS", "Inputs", f"Equilibrium column is '{assoc[0]}'", f"association time {t_assoc:g} s")

    # --- raw files: which file holds which ligand and concentration comes from the results table
    try:
        files_by = R.ligand_files(rows, ligands, series)
    except ValueError as e:
        ck.add("FAIL", "Inputs", "Cannot tell which raw files belong to which ligand and concentration", str(e))
        return res
    res["files"] = files_by
    legacy = all(set(files_by[l]) == set(range(len(series))) and
                 files_by[l][0] == f"A{pos}.xls" for pos, l in enumerate(ligands, 1))
    ck.add("PASS" if legacy else "INFO", "Inputs",
           "Raw files, from the results table's Sensor Location: " + "; ".join(
               f"{l} = {R.describe_files(files_by[l])}" for l in ligands),
           "each ligand's files at each concentration are found by sensor location, so several ligands "
           "may share a sensor column and a ligand may have fewer than 8 concentrations")
    for l in ligands:
        have = len(files_by[l])
        if have < n_conc:
            ck.add("INFO", "Inputs", f"{l} has {have} concentration(s), 'Kinetic concentrations per ligand' is {n_conc}",
                   f"only the {have} that exist are built")
        if have < len(series):
            ck.add("INFO", "Inputs", f"{l} has {have} of the Setup series' {len(series)} dilution steps",
                   "the missing steps are simply absent from the results table")
    raw, file_of = {}, {}
    for i, lig in enumerate(ligands, 1):
        present = sorted(files_by[lig])
        for rank, step in enumerate(present):
            tag = files_by[lig][step]
            file_of[(i, step)] = tag
            needed = rank < n_conc
            path = os.path.join(folder, tag)
            if not os.path.exists(path):
                ck.add("FAIL" if needed else "INFO", "Raw files", f"{tag} is missing",
                       f"{lig}, {series[step] * 1e9:.4g} nM: " + ("needed for a kinetics sheet" if needed
                                                               else "not needed (beyond the kinetic concentration count)"))
                continue
            raw[(i, step)] = R.parse_raw_xls(path)
    res["raw"] = raw
    bad_conc = bad_shape = 0
    grids = {}
    for (i, step), r in raw.items():
        tag = file_of[(i, step)]
        if len(r["data"]) != len(analytes):
            bad_shape += 1
            ck.add("FAIL", "Raw files", f"{tag} has {len(r['data'])} traces, setup has {len(analytes)} analytes")
        hdr = r["concs"]
        if max(hdr) / min(hdr) > 1.0001:
            bad_conc += 1
            ck.add("FAIL", "Raw files", f"{tag}: traces disagree on Conc1", str(hdr))
        else:
            try:
                idx = R.match_step(hdr[0], series)
                if idx != step:
                    bad_conc += 1
                    ck.add("FAIL", "Raw files", f"{tag} is dilution step {idx + 1}, but the results table puts it at step {step + 1}",
                           f"Conc1 = {hdr[0]:g} M")
            except ValueError as e:
                bad_conc += 1
                ck.add("FAIL", "Raw files", f"{tag}: Conc1 matches no Setup dilution step", str(e))
        t = _floats(r["times"])
        grids[(i, step)] = (len(t), float(t[0]), float(t[-1]), round(float(np.median(np.diff(t))), 6))
        blanks = sum(int(np.isnan(_floats(col)).sum()) for col in r["data"])
        if blanks:
            ck.add("FAIL", "Raw files", f"{tag} has {blanks} blank data values")
    if raw and bad_conc == 0 and bad_shape == 0:
        ck.add("PASS", "Raw files", f"{len(raw)} raw file(s) match the setup",
               "trace count equals analyte count and every Conc1 maps to its dilution step")
    if len(set(grids.values())) > 1:
        ck.add("FAIL", "Raw files", "Files do not share one time grid", str(sorted(set(grids.values()))))
    elif grids:
        n, t0, t1, dt = next(iter(grids.values()))
        ck.add("PASS", "Raw files", "All files share one time grid", f"{n} points, {t0:g} to {t1:g} s, {dt:g} s apart")
    if raw and (devs := [abs(r["concs"][0] / series[s] - 1) for (i, s), r in raw.items()]):
        worst = max(devs) * 100
        ck.add("WARN" if worst > 5 else "INFO", "Raw files", "Conc1 headers are rounded",
               f"largest deviation from the exact Setup dilution step is {worst:.2f}% (Octet rounds Conc1 to "
               "as little as one significant figure); the tool uses the exact Setup value for the X axis. "
               "Files are matched to a step within 10%." + (" This deviation is large: check the dilution series."
                                                           if worst > 5 else ""))

    # --- traces that fall during association
    declining = []
    for (i, step), r in sorted(raw.items()):
        t = _floats(r["times"])
        inside = t <= t_assoc + 1e-9
        for j, a in enumerate(analytes[:len(r["data"])]):
            y = _floats(r["data"][j])
            sm = np.convolve(y[inside], np.ones(5) / 5, mode="valid")
            peak, end = float(sm.max()), float(np.mean(y[inside][-10:]))
            if peak > ASSOC_MIN_SIGNAL_NM and end < (1 - ASSOC_DECLINE) * peak:
                declining.append((ligands[i - 1], a["display"], R.conc_label_nM(series[step]), peak, end))
    res["declining"] = declining
    if declining:
        ck.add("WARN", "Raw traces", f"{len(declining)} trace(s) fall during the association phase",
               "; ".join(f"{l} / {a} at {c}: peak {p:.3f} nm, {e:.3f} nm at the end ({(1 - e / p) * 100:.0f}% lower)"
                         for l, a, c, p, e in declining)
               + ". A response that rises then falls while analyte is present is not simple 1:1 binding "
               "(aggregation, bulk effect, or a second process), so kinetic fits of these traces are unreliable.")
    elif raw:
        ck.add("PASS", "Raw traces", "No trace falls during the association phase")

    # --- CSV completeness and CSV vs raw
    try:
        eq = {}
        for i, lig in enumerate(ligands, 1):
            present = sorted(files_by[lig])
            by_step, _ = R.step_concentrations(folder, files_by[lig], series, True)
            xs = [by_step[s_] for s_ in present]
            eq[lig] = (xs, [[float(v) for v in ln.split(",")[1:]]
                            for ln in R.equilibrium_table(rows, lig, series, xs, analytes, steps=present)])
        res["eq"] = eq
        ck.add("PASS", "Results CSV", "Every ligand × analyte × concentration has exactly one row",
               f"{sum(len(v[1]) * len(analytes) for v in eq.values())} equilibrium values")
    except ValueError as e:
        ck.add("FAIL", "Results CSV", "CSV rows do not line up with the setup", str(e))
        return res

    cmp_rows = []
    for (i, step), r in raw.items():
        t = _floats(r["times"])
        k = int(np.argmin(np.abs(t - t_assoc)))
        if abs(t[k] - t_assoc) > 1e-6:
            ck.add("FAIL", "Results CSV", f"{file_of[(i, step)]} has no sample at t = {t_assoc:g} s")
            continue
        for j, a in enumerate(analytes):
            if j >= len(r["data"]):
                continue
            raw_v = float(r["data"][j][k])
            csv_v = eq[ligands[i - 1]][1][sorted(files_by[ligands[i - 1]]).index(step)][j]
            cmp_rows.append((ligands[i - 1], a["display"], R.conc_label_nM(series[step]), csv_v, raw_v, csv_v - raw_v))
    res["cmp"] = cmp_rows
    if cmp_rows:
        d = np.array([c[5] for c in cmp_rows])
        over = [c for c in cmp_rows if abs(c[5]) > CSV_RAW_WARN_NM]
        detail = (f"{len(cmp_rows)} comparisons; median |difference| {np.median(np.abs(d)):.5f} nm, "
                  f"largest {np.abs(d).max():.5f} nm. The CSV value is the one used (design decision); "
                  "the raw trace is processed slightly differently by Octet.")
        if over:
            ck.add("WARN", "Results CSV", f"{len(over)} equilibrium value(s) differ from the raw trace by more than {CSV_RAW_WARN_NM} nm", detail)
        else:
            ck.add("PASS", "Results CSV", f"CSV equilibrium values agree with the raw traces at t = {t_assoc:g} s", detail)

    # --- the top concentration should be the highest response
    for lig in ligands:
        vals = eq[lig][1]
        for j, a in enumerate(analytes):
            col = [v[j] for v in vals]
            if col[0] > 0 and max(col) > HOOK_FACTOR * col[0]:
                ck.add("WARN", "Equilibrium values", f"{lig} / {a['display']}: the response is not highest at the top concentration",
                       f"the top concentration gives {col[0]:.3f} nm but the highest response is {max(col):.3f} nm "
                       f"({max(col) / col[0] * 100:.0f}% of it): a hook effect, aggregation or a bad trace. The "
                       "workbook's 'normalized to 100' columns divide by the top-concentration value, so they exceed "
                       "100, and a sigmoid fit through the top point will be poor. Consider excluding that point in "
                       "the Prism fit.")
    # --- equilibrium fits
    fits = {}
    for lig in ligands:
        xs, vals = eq[lig]
        x = np.log10(np.array(xs))
        for j, a in enumerate(analytes):
            y = np.array([v[j] for v in vals])
            f = eqfit.fit_curve(x, y)
            fits[(lig, a["csv_id"])] = f
            problems = eqfit.flags(f)
            if problems:
                ck.add("WARN", "Equilibrium fits", f"{lig} / {a['display']}", "; ".join(problems))
    res["fits"] = fits
    n_bad = sum(1 for (l, a), f in fits.items() if eqfit.flags(f))
    if n_bad == 0:
        ck.add("PASS", "Equilibrium fits", "All equilibrium fits are well behaved")
    else:
        ck.add("INFO", "Equilibrium fits", f"{n_bad} of {len(fits)} equilibrium fits have a warning",
               "see the individual warnings and the fit table")

    # --- kinetic fits
    fit_rows = K.read_fit_rows(results_csv)
    table = K.build_table(fit_rows, ligands, names)
    res["ktable"] = table
    groups = {}
    for lig in ligands:
        for a in analytes:
            g = [c for c in table if c["ligand"] == lig and c["analyte"] == a["csv_id"]]
            used = [c for c in g if c["status"] == "used"]
            row = {"ligand": lig, "analyte": a["display"], "n": len(g), "used": len(used),
                   "flagged": sum(c["flagged"] for c in used),
                   "kd2_only": sum(1 for c in g if "second site" in c["reason"])}
            if used:
                kon = statistics.mean(c["kon"] for c in used)
                kdis = statistics.mean(c["kdis"] for c in used)
                row.update(kd_means=kdis / kon * 1e9, kd_rows=statistics.mean(c["kd_M"] for c in used) * 1e9)
                row["ratio"] = max(row["kd_means"], row["kd_rows"]) / min(row["kd_means"], row["kd_rows"])
            groups[(lig, a["csv_id"])] = row
            where = f"{lig} / {a['display']}"
            if not used:
                extra = (f" ({row['kd2_only']} row(s) have only a 2:1 second-site KD)" if row["kd2_only"] else "")
                ck.add("WARN", "Kinetic fits", f"{where}: no fit row passed the filters", f"no KD is reported for this pair{extra}")
            else:
                if row["flagged"]:
                    ck.add("WARN", "Kinetic fits", f"{where}: {row['flagged']} flagged row(s) included",
                           "; ".join(f"{c['conc_nM']:g} nM: {c['reason']}" for c in used if c["flagged"]))
                if row["ratio"] > KD_DISAGREE_FOLD:
                    ck.add("WARN", "Kinetic fits", f"{where}: the two KD estimates differ {row['ratio']:.1f}-fold",
                           f"KD from mean kdis / mean kon = {row['kd_means']:.3g} nM; mean of per-row KDs = {row['kd_rows']:.3g} nM")
                if len(used) == 1:
                    ck.add("INFO", "Kinetic fits", f"{where}: a single fit row", "no replication, so no spread to judge")
    res["kgroups"] = groups
    ck.add("INFO", "Kinetic fits", f"{sum(1 for c in table if c['status'] == 'used')} of {len(table)} fit rows used",
           "see the exclusion table")

    # --- Prism project dry run
    if template or scratch:
        try:
            from .rebuild import build
            scratch = scratch or template is None
            proj, problems, _, labels = build(folder, template, None, n_conc, setup_path, scratch=scratch, container=container)
            res["prism"] = {"problems": problems, "warnings": proj.warnings, "labels": labels}
            if problems:
                ck.add("FAIL", "Prism project", f"{len(problems)} structural problem(s) in the project", "; ".join(problems[:5]))
            else:
                what = ("Project built from scratch validates" if scratch else "Reshaped project validates")
                ck.add("PASS", "Prism project", what,
                       f"{len(ligands)} ligand(s) × {len(labels)} concentration(s): every sheet, dataset and analysis reference resolves"
                       + ("; no graphs or layouts are written, create them in Prism" if scratch else ""))
            for w in proj.warnings:
                if scratch:
                    ck.add("INFO", "Prism project", "No graphs or layouts", w)
                else:
                    ck.add("WARN", "Prism project", "Layout removed", w)
        except Exception as e:  # ShapeError and friends are reportable, not fatal
            ck.add("FAIL", "Prism project", "Template cannot be reshaped to this run", f"{type(e).__name__}: {e}")
    else:
        ck.add("INFO", "Prism project", "Project checks skipped")
    return res


# ----------------------------------------------------------------------------- figures
def fig_equilibrium(res, lig):
    xs, vals = res["eq"][lig]
    x = np.log10(np.array(xs))
    xx = np.linspace(x.min() - 0.3, x.max() + 0.3, 300)
    fig, ax = plt.subplots(figsize=(6.4, 4))
    colors = plt.cm.tab10.colors
    for j, a in enumerate(res["analytes"]):
        y = np.array([v[j] for v in vals])
        ax.plot(x, y, "o", color=colors[j], label=a["display"])
        f = res["fits"][(lig, a["csv_id"])]["three"]
        if f["ok"]:
            ax.plot(xx, eqfit.sigmoid3(xx, *f["params"]), "-", color=colors[j], alpha=0.8)
    ax.set_xlabel("log10 [analyte] (M)")
    ax.set_ylabel("response at t = end of association (nm)")
    ax.set_title(f"{lig}: equilibrium binding, 3-parameter fit")
    ax.legend(fontsize=8, frameon=False)
    ax.grid(alpha=0.25)
    return _png(fig)


def fig_sensorgrams(res, i, lig):
    steps = sorted(s for (ii, s) in res["raw"] if ii == i)
    if not steps:
        return None
    series = res["setup"]["series_M"]
    n = len(res["analytes"])
    ncols = min(n, 4)
    nrows = -(-n // ncols)
    fig, axes = plt.subplots(nrows, ncols, figsize=(4.2 * ncols, 3.4 * nrows), sharey=True, squeeze=False)
    flat = axes.flatten()
    for extra in flat[n:]:
        extra.axis("off")
    axes = flat[:n]
    cmap = plt.cm.viridis_r
    for j, (ax, a) in enumerate(zip(axes, res["analytes"])):
        for s in steps:
            r = res["raw"][(i, s)]
            t, y = _floats(r["times"]), _floats(r["data"][j])
            ax.plot(t[::4], y[::4], lw=0.9, color=cmap(s / max(len(series) - 1, 1)),
                    label=R.conc_label_nM(series[s]))
        ax.axvline(float(res.get("t_assoc", 295)), color="grey", ls=":", lw=0.8)
        ax.set_title(a["display"], fontsize=9)
        ax.set_xlabel("time (s)")
        ax.grid(alpha=0.25)
    for row in axes.reshape(-1)[::ncols]:
        row.set_ylabel("response (nm)")
    axes[-1].legend(fontsize=6, frameon=False, ncol=2, title="conc", title_fontsize=6)
    fig.suptitle(f"{lig}: sensorgrams (dotted line = end of association)", fontsize=10, y=1.02)
    return _png(fig)


# ------------------------------------------------------------------------------- html
CSS = """
body{font-family:-apple-system,Helvetica,Arial,sans-serif;max-width:1100px;margin:2em auto;padding:0 1em;color:#222}
h1{margin-bottom:.2em}h2{margin-top:2em;border-bottom:2px solid #1f3864;padding-bottom:.2em;color:#1f3864}
table{border-collapse:collapse;margin:.8em 0;font-size:.88em;display:block;overflow-x:auto;max-width:100%}th,td{border:1px solid #ccc;padding:4px 8px;text-align:left;vertical-align:top}
th{background:#dce6f1}td.n{text-align:right;font-variant-numeric:tabular-nums}
.FAIL{background:#f8d7da}.WARN{background:#fff3cd}.PASS{background:#d4edda}.INFO{background:#e8eef7}
.badge{display:inline-block;padding:.15em .6em;border-radius:3px;font-weight:600;margin-right:.5em}
.muted{color:#666;font-size:.9em}img{max-width:100%}code{background:#f2f2f2;padding:0 3px}
"""


def _e(x):
    return html.escape(str(x))


def _fmt(v, spec):
    return format(v, spec) if isinstance(v, (int, float)) and np.isfinite(v) else "n/a"


def render(res, title="bli2prism QC report"):
    ck = res["checks"]
    out = [f"<!doctype html><html><head><meta charset='utf-8'><title>{_e(title)}</title><style>{CSS}</style></head><body>"]
    out.append(f"<h1>{_e(title)}</h1>")
    setup = res.get("setup")
    name = setup.get("name") if isinstance(setup, dict) else None
    out.append(f"<p class='muted'>Folder: <code>{_e(os.path.basename(os.path.abspath(res['folder'])))}</code>"
               + (f" &middot; experiment: {_e(name)}" if name else "") + f" &middot; kinetic concentrations: {res['n_conc']}</p>")
    out.append("<p>" + "".join(f"<span class='badge {l}'>{ck.count(l)} {l}</span>" for l in ("FAIL", "WARN", "PASS", "INFO")) + "</p>")
    verdict = ("<b>Do not use these outputs until the FAIL items are fixed.</b>" if ck.count("FAIL")
               else "No blocking problems. Read the warnings before using the numbers." if ck.count("WARN")
               else "All checks passed.")
    out.append(f"<p>{verdict}</p>")

    out.append("<h2>Checks</h2><table><tr><th>Result</th><th>Area</th><th>Check</th><th>Detail</th></tr>")
    for c in sorted(ck.items, key=lambda c: LEVELS[c["level"]]):
        out.append(f"<tr class='{c['level']}'><td>{c['level']}</td><td>{_e(c['area'])}</td>"
                   f"<td>{_e(c['title'])}</td><td>{_e(c['detail'])}</td></tr>")
    out.append("</table>")

    if "fits" in res:
        out.append("<h2>Equilibrium fits</h2><p class='muted'>Sigmoidal dose-response, fixed slope, fitted here in Python "
                   "to the CSV equilibrium values; X = log10(M). The 4-parameter column adds a Hill slope for comparison "
                   "(the Prism project applies the fixed-slope fit). Intervals are approximate (Wald); Prism reports "
                   "profile-likelihood intervals.</p>")
        for lig in res["setup"]["ligands"]:
            out.append(f"<h3>{_e(lig)}</h3><img src='data:image/png;base64,{fig_equilibrium(res, lig)}'>")
            out.append("<table><tr><th>Analyte</th><th>Bottom</th><th>Top</th><th>LogEC50 (95% CI)</th><th>EC50 (nM)</th>"
                       "<th>R²</th><th>Sy.x</th><th>Hill slope (4P)</th><th>ΔAICc (3P − 4P)</th><th>Warnings</th></tr>")
            for a in res["analytes"]:
                f = res["fits"][(lig, a["csv_id"])]
                t, q = f["three"], f["four"]
                if not t["ok"]:
                    out.append(f"<tr class='FAIL'><td>{_e(a['display'])}</td><td colspan='9'>{_e(t.get('error'))}</td></tr>")
                    continue
                b, top, lec = t["params"]
                lo, hi = t["ci"][2]
                dA = (t["aicc"] - q["aicc"]) if q["ok"] and np.isfinite(t["aicc"]) and np.isfinite(q["aicc"]) else float("nan")
                fl = eqfit.flags(f)
                out.append(f"<tr class='{'WARN' if fl else ''}'><td>{_e(a['display'])}</td><td class='n'>{b:.4f}</td>"
                           f"<td class='n'>{top:.4f}</td><td class='n'>{lec:.3f} ({_fmt(lo, '.3f')} to {_fmt(hi, '.3f')})</td>"
                           f"<td class='n'>{10 ** lec * 1e9:.3g}</td><td class='n'>{t['r2']:.4f}</td><td class='n'>{t['syx']:.4f}</td>"
                           f"<td class='n'>{_fmt(q['params'][3], '.2f') if q['ok'] else 'n/a'}</td>"
                           f"<td class='n'>{_fmt(dA, '.1f')}</td><td>{_e('; '.join(fl))}</td></tr>")
            out.append("</table>")

    if res.get("cmp"):
        out.append("<h2>Equilibrium values: results CSV vs raw trace</h2><p class='muted'>The CSV <code>X=295</code> value is used; "
                   "this table shows how far the raw trace at the same time is from it. Rows over "
                   f"{CSV_RAW_WARN_NM} nm are highlighted; all rows listed.</p><table><tr><th>Ligand</th><th>Analyte</th>"
                   "<th>Conc</th><th>CSV (nm)</th><th>Raw (nm)</th><th>Difference (nm)</th></tr>")
        for lig, an, conc, c, r, d in sorted(res["cmp"], key=lambda x: -abs(x[5])):
            cls = "WARN" if abs(d) > CSV_RAW_WARN_NM else ""
            out.append(f"<tr class='{cls}'><td>{_e(lig)}</td><td>{_e(an)}</td><td>{_e(conc)}</td><td class='n'>{c:.5f}</td>"
                       f"<td class='n'>{r:.5f}</td><td class='n'>{d:+.5f}</td></tr>")
        out.append("</table>")

    if "kgroups" in res:
        out.append("<h2>Kinetic fits</h2><p class='muted'>Rows used: Full X² &lt; 3, Full R² &gt; 0.9, and KD, kon, kdis all "
                   "determinate. Flagged = used but KD more than 5-fold from the median of the group's used rows "
                   "(still included). The Excel workbook lists every row.</p>")
        out.append("<table><tr><th>Ligand</th><th>Analyte</th><th>Rows</th><th>Used</th><th>Flagged</th>"
                   "<th>KD from means (nM)</th><th>Mean of row KDs (nM)</th><th>Ratio</th><th>2:1-only rows</th></tr>")
        for g in res["kgroups"].values():
            cls = "WARN" if (g["used"] == 0 or g["flagged"] or g.get("ratio", 1) > KD_DISAGREE_FOLD) else ""
            out.append(f"<tr class='{cls}'><td>{_e(g['ligand'])}</td><td>{_e(g['analyte'])}</td><td class='n'>{g['n']}</td>"
                       f"<td class='n'>{g['used']}</td><td class='n'>{g['flagged']}</td>"
                       f"<td class='n'>{_fmt(g.get('kd_means'), '.3g')}</td><td class='n'>{_fmt(g.get('kd_rows'), '.3g')}</td>"
                       f"<td class='n'>{_fmt(g.get('ratio'), '.2f')}</td><td class='n'>{g['kd2_only']}</td></tr>")
        out.append("</table>")
        reasons = {}
        for c in res["ktable"]:
            if c["status"] == "excluded":
                key = ("KD censored (second-site KD available)" if "second site" in c["reason"]
                       else "KD censored" if c["reason"].startswith("KD censored")
                       else "Full R² too low" if "R²" in c["reason"]
                       else "Full X² too high" if "X²" in c["reason"]
                       else "kon did not converge" if c["reason"].startswith("kon")
                       else "kdis censored" if c["reason"].startswith("kdis") else c["reason"])
                reasons[key] = reasons.get(key, 0) + 1
        out.append("<table><tr><th>Why fit rows were excluded</th><th>Rows</th></tr>"
                   + "".join(f"<tr><td>{_e(k)}</td><td class='n'>{v}</td></tr>" for k, v in sorted(reasons.items(), key=lambda kv: -kv[1]))
                   + "</table>")

    if res.get("raw"):
        out.append("<h2>Sensorgrams</h2><p class='muted'>All available concentrations. Look for traces that do not "
                   "rise with concentration, drift, or noise that swamps the signal.</p>")
        for i, lig in enumerate(res["setup"]["ligands"], 1):
            img = fig_sensorgrams(res, i, lig)
            if img:
                out.append(f"<img src='data:image/png;base64,{img}'>")

    if "prism" in res:
        p = res["prism"]
        out.append("<h2>Prism project</h2>")
        out.append(f"<p>Concentrations per ligand: {_e(', '.join(p['labels']))}</p>")
        out.append("<p>" + (_e("; ".join(p["problems"])) if p["problems"] else "Structural validation found no problems.") + "</p>")
        for w in p["warnings"]:
            out.append(f"<p class='muted'>Note: {_e(w)}</p>")

    out.append("<h2>Not covered</h2><p class='muted'>This report cannot tell whether the reference subtraction was correct "
               "(exports are assumed to be double-referenced), whether the Prism file opens (open it once), or whether the "
               "chosen kinetic model is appropriate for the interaction.</p></body></html>")
    return "".join(out)


def write_report(res, out_path):
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(render(res))
    return out_path
