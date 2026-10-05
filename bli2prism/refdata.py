"""Data layer for the reference rebuild: raw Octet exports, results CSV, Setup sheet
-> the tables that go into the Prism project.

Everything here returns strings exactly as they should appear in a table's data.csv,
so that a rebuild from the same inputs can be compared byte-for-byte with the original.
"""

from __future__ import annotations

import csv
import math
import os
import re

def parse_raw_xls(path):
    """Octet 'xls' export (tab-delimited text) -> {conc_M: [floats], times: [str], data: [[str]]}.

    Values stay as the original text so they can be written back unchanged.
    """
    with open(path, encoding="latin1", newline="") as fh:
        lines = fh.read().replace("\r\n", "\n").split("\n")
    concs = [float(x) for x in lines[1].split("\t")[1:] if x != ""]
    names = lines[4].rstrip("\t").split("\t")
    n_traces = len(names) // 3
    times, data = [], [[] for _ in range(n_traces)]
    for ln in lines[5:]:
        if not ln.strip():
            continue
        cells = ln.rstrip("\t").split("\t")
        # Time<i>, Data<i>, Sim<i> triplets; trailing empty Sim cells may be dropped
        cells += [""] * (3 * n_traces - len(cells))
        times.append(cells[0])
        for i in range(n_traces):
            data[i].append(cells[3 * i + 1])
    return {"concs": concs, "times": times, "data": data}


def read_results_csv(path):
    with open(path, newline="") as fh:
        return list(csv.DictReader(fh))


def read_setup(xlsx_path, loading_ids):
    """What the pipeline needs from the setup workbook (its bli2prism sheet): ligands, reference, analytes with their
    molecular weights, the dilution series and the manual inputs. loading_ids: the 'Loading Sample ID' values found in
    the results CSV, which must be exactly the sheet's ligands."""
    from . import setupsheet

    if not setupsheet.has_sheet(xlsx_path):
        raise ValueError(f"{xlsx_path} has no bli2prism input sheet: run `bli2prism draft` to create one")
    s = setupsheet.read(xlsx_path)
    if set(s["ligands"]) != set(loading_ids):
        raise ValueError(setupsheet.explain_ligand_mismatch(s, loading_ids))
    return s


def log_x(conc_M):
    return f"{math.log10(conc_M):.9f}".rstrip("0").rstrip(".")


def match_step(header_M, series_M, rel_tol=0.10):
    """Index of the Setup dilution step a file's (rounded) Conc1 header belongs to.

    Octet rounds Conc1 to as few as one significant figure (3E-10 for 3.05E-10, a 1.7% error),
    and adjacent steps differ by the dilution factor (>= 2x), so 10% is safe.
    """
    best = min(range(len(series_M)), key=lambda i: abs(math.log(header_M / series_M[i])))
    if abs(header_M / series_M[best] - 1) > rel_tol:
        raise ValueError(f"Conc1 {header_M:g} M matches no Setup dilution step "
                         f"(nearest {series_M[best]:g} M)")
    return best


def conc_label_nM(conc_M):
    """Sheet-title rule: truncate to integer nM; sub-nM to 3 significant figures."""
    nM = conc_M * 1e9
    if nM >= 1:
        return f"{int(nM + 1e-9)} nM"
    return f"{float(f'{nM:.3g}'):g} nM"


def step_concentrations(folder, files, series_M, snap=True):
    """Concentration (M) of each dilution step a ligand has, read from its raw files' Conc1 header.

    files: {step: raw file name} for the steps the ligand has (see ligand_files). Returns
    ({step: concentration in M}, warnings).

    Octet writes Conc1 rounded to as little as one significant figure (7.81E-08 for 7.8125E-08).
    With snap=True a header within 10% of a Setup dilution step is replaced by that step's exact
    value, which is what the reference project's X axis uses; with no match the raw header is
    kept and a warning is returned alongside. A missing file falls back to the Setup value.
    """
    out, warnings = {}, []
    for step, name in sorted(files.items()):
        path = os.path.join(folder, name)
        if not os.path.exists(path):
            out[step] = series_M[step]
            warnings.append(f"{name} is missing; using the Setup value for step {step + 1}")
            continue
        hdr = parse_raw_xls(path)["concs"][0]
        try:
            idx = match_step(hdr, series_M)
            out[step] = series_M[idx] if snap else hdr
        except ValueError as e:
            out[step] = hdr
            warnings.append(f"{name}: {e}; using raw header")
    return out, warnings


def equilibrium_table(rows, ligand, series_M, xs, analytes, steps=None):
    """-> list of csv lines: log10(M), then one response column per analyte (X=295 column).

    xs: concentration (M) of each row for the X column; defaults to series_M.
    steps: the dilution step index of each row (default 0, 1, 2, ...): a ligand may have only
    some steps, e.g. the lower four of a series.
    """
    xs = series_M if xs is None else xs
    steps = list(range(len(xs))) if steps is None else steps
    assoc_col = [k for k in rows[0] if k.startswith("X=") and k != "X=0"]
    if len(assoc_col) != 1:
        raise ValueError(f"expected one non-zero X=<t> column, found {assoc_col}")
    col = assoc_col[0]
    out = []
    for step, c in zip(steps, xs):
        line = [log_x(c)]
        for a in analytes:
            sel = [r for r in rows if r["Loading Sample ID"] == ligand
                   and r["Sample ID"] == a["csv_id"]
                   and match_step(float(r["Conc. (nM)"]) * 1e-9, series_M, 0.05) == step]
            if len(sel) != 1:
                raise ValueError(f"{ligand}/{a['csv_id']}/step {step}: {len(sel)} rows")
            line.append(sel[0][col].strip())
        out.append(",".join(line))
    return out


def kinetics_table(raw, analytes):
    """-> list of csv lines: time, then one trace per analyte, original text."""
    n = len(raw["times"])
    n_traces = len(analytes)
    if len(raw["data"]) < n_traces:
        raise ValueError(f"raw file has {len(raw['data'])} traces but {n_traces} analytes are set up")
    return [",".join([raw["times"][i]] + [raw["data"][k][i] for k in range(n_traces)])
            for i in range(n)]


def ligand_files(rows, ligands, series_M, rel_tol=0.05):
    """Which raw file holds each ligand's data at each dilution step: {ligand: {step: 'E2.xls'}}.

    Read from the results table: every row has a 'Sensor Location' ('E2': row E, sensor column 2)
    and a concentration, and the file is named after the location. Reading it rather than
    assuming a layout handles every arrangement seen so far: one ligand per sensor column with
    all 8 rows (the reference), a single ligand on column 5, and several ligands stacked on one
    column (rows A-D and E-H). A step is the position of the concentration in the Setup dilution
    series. If the CSV has no Sensor Location, the old convention is used: letter = step, number
    = the ligand's position.
    """
    out = {}
    for pos, lig in enumerate(ligands, 1):
        by_step = {}
        for r in rows:
            if r.get("Loading Sample ID") != lig:
                continue
            try:
                step = match_step(float(r["Conc. (nM)"]) * 1e-9, series_M, rel_tol)
            except ValueError:
                raise ValueError(f"{lig}: the results table has {r['Conc. (nM)']} nM, which matches no step of "
                                 "the Setup dilution series (check the dilution start, factor and steps)") from None
            loc = (r.get("Sensor Location") or "").strip()
            by_step.setdefault(step, set()).add(loc or f"{'ABCDEFGH'[step]}{pos}")
        for step, locs in by_step.items():
            if len(locs) > 1:
                raise ValueError(f"{lig} has several sensors {sorted(locs)} at {series_M[step] * 1e9:.4g} nM; "
                                 "one sensor per ligand and concentration is supported")
        out[lig] = {step: f"{next(iter(locs))}.xls" for step, locs in sorted(by_step.items())}
    used = {}
    for lig, d in out.items():
        for name in d.values():
            if used.setdefault(name, lig) != lig:
                raise ValueError(f"{name} is claimed by both {used[name]!r} and {lig!r}: check the results "
                                 "table's 'Sensor Location' column")
    return out


def describe_files(files):
    """'A2-D2' style summary of one ligand's files, for messages."""
    names = [f[:-4] for _, f in sorted(files.items())]
    return names[0] if len(names) == 1 else f"{names[0]}-{names[-1]}" if names else "(none)"


