"""Kinetic fit-row selection from the Octet results CSV.

Every fit row gets a status and a reason, so nothing is dropped silently. Rules (design
decisions kinetic_row_selection, averaging_convention, kinetic_kd_loose_ends):

    used      Full X^2 < 3 and Full R^2 > 0.9 and KD, kon, kdis all determinate
    flagged   used rows whose KD is more than 5-fold from the median KD of the group's used
              rows (needs >= 3 used rows). The median resists outliers up to 50% of the rows, so
              a few bad rows cannot make the good ones look bad. Flagged rows stay included.
    excluded  everything else, with the first failing reason

A fit is not determinate when Octet reports a censored bound ('<1.0E-12', '<1.0E-07') or a
non-converged value (kon = 4.2E+99).
"""

from __future__ import annotations

import csv
import statistics

X2_MAX = 3.0
R2_MIN = 0.9
OUTLIER_FOLD = 5.0
NONCONVERGED = 1e30          # kon/kdis at or above this are fit failures, not measurements


def read_fit_rows(path):
    """Results CSV -> list of dicts keyed by header, taking the FIRST of any duplicate header
    (the CSV repeats 'Full  X^2' / 'Full  R^2')."""
    with open(path, newline="") as fh:
        rd = csv.reader(fh)
        header = next(rd)
        first = {}
        for i, h in enumerate(header):
            first.setdefault(h, i)
        return [{h: row[i] for h, i in first.items() if i < len(row)} for row in rd]


def _num(text):
    """float if the cell is a determinate number, else None."""
    t = (text or "").strip()
    if not t or t[0] in "<>":
        return None
    try:
        v = float(t)
    except ValueError:
        return None
    return v


def classify(row):
    """-> dict(kd_M, kon, kdis, x2, r2, status, reason) for one CSV row.

    Only the `KD (M)` / `kon(1/Ms)` / `kdis(1/s)` columns are ever used. The 2:1 second-site
    columns (KD2 ...) are deliberately ignored (design decision: KD values only); a row whose
    KD is censored is excluded even if its KD2 is determinate, and its reason says so.
    """
    x2, r2 = _num(row.get("Full  X^2")), _num(row.get("Full  R^2"))
    out = {"x2": x2, "r2": r2, "kd_M": _num(row.get("KD (M)")),
           "kon": _num(row.get("kon(1/Ms)")), "kdis": _num(row.get("kdis(1/s)")),
           "kd_text": row.get("KD (M)", ""), "kon_text": row.get("kon(1/Ms)", ""),
           "kdis_text": row.get("kdis(1/s)", "")}
    kd2 = _num(row.get("KD2"))
    if x2 is None or x2 >= X2_MAX:
        out.update(status="excluded", reason=f"Full X² {row.get('Full  X^2')} is not < {X2_MAX:g}")
        return out
    if r2 is None or r2 <= R2_MIN:
        out.update(status="excluded", reason=f"Full R² {row.get('Full  R^2')} is not > {R2_MIN:g}")
        return out
    if out["kd_M"] is None:
        why = f"KD censored ({out['kd_text']})"
        if kd2 is not None:
            why += f"; only the 2:1 second site is determinate (KD2 = {kd2:.3g} M)"
        out.update(status="excluded", reason=why)
        return out
    if out["kon"] is None or out["kon"] >= NONCONVERGED:
        out.update(status="excluded", reason=f"kon did not converge ({out['kon_text']})")
        return out
    if out["kdis"] is None or out["kdis"] >= NONCONVERGED:
        out.update(status="excluded", reason=f"kdis censored ({out['kdis_text']})")
        return out
    out.update(status="used", reason="")
    return out


def build_table(rows, ligands, analytes):
    """-> list of dict, one per fit row, grouped ligand -> analyte -> CSV order, with flags."""
    out = []
    for lig in ligands:
        for a in analytes:
            grp = [r for r in rows if r["Loading Sample ID"] == lig and r["Sample ID"] == a]
            recs = []
            for r in grp:
                c = classify(r)
                c.update(ligand=lig, analyte=a, conc_nM=float(r["Conc. (nM)"]), flagged=False)
                recs.append(c)
            used = [c for c in recs if c["status"] == "used"]
            if len(used) >= 3:
                med = statistics.median(c["kd_M"] for c in used)
                for c in used:
                    fold = max(c["kd_M"] / med, med / c["kd_M"])
                    if fold > OUTLIER_FOLD:
                        c["flagged"] = True
                        c["reason"] = (f"flagged: KD is {fold:.1f}-fold from the median of the "
                                       f"{len(used)} used rows ({med * 1e9:.3g} nM); still included")
            out += recs
    return out
