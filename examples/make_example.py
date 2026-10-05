"""Generate the demo experiment's raw files and results table from its setup workbook, using nothing but random numbers.

No real data is involved: the traces are simulated 1:1 binding curves with seeded noise, and the results table is produced
by fitting the simulated traces. The ligands, analytes and dilution series are read from
examples/demo_binder_screen/bli2prism_setup.xlsx, which is the input to this script (edit the plate map there, then rerun).

    python3 examples/make_example.py                 # rewrites the files in examples/demo_binder_screen/
    python3 examples/make_example.py FOLDER          # a folder that already holds bli2prism_setup.xlsx

What it writes (the same files a real Octet run folder holds):
    A1.xls .. H1.xls      sensor column 1 = the first ligand, plate rows A-H = the 8 analyte concentrations
    A2.xls .. H2.xls      sensor column 2 = the second ligand
    kineticanalysistableresults.csv   one fit row per ligand / analyte / concentration
"""

from __future__ import annotations

import csv
import math
import os
import sys

import numpy as np
from scipy.optimize import curve_fit

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from bli2prism import setupsheet as S  # noqa: E402

SEED = 20260101
T_ASSOC, T_END, DT = 300.0, 600.0, 0.5
RMAX = 0.55                                               # nm
# (kon 1/Ms, kdis 1/s) by (ligand number, analyte number) in the setup's order; None = no binding (drift and noise)
KINETICS = {
    (1, 1): (2.0e5, 1.0e-3),       # KD  5 nM
    (1, 2): (4.0e5, 4.0e-4),       # KD  1 nM
    (1, 3): None,                  # does not bind
    (2, 1): (1.0e5, 5.0e-3),       # KD 50 nM
    (2, 2): (2.0e5, 4.0e-3),       # KD 20 nM
    (2, 3): (2.0e4, 4.0e-2),       # KD  2 uM (weak)
}
NOISE = 0.002


def simulate(rng, kin, conc_M):
    """Reference-subtracted response (nm) over the whole run: association to T_ASSOC, then dissociation."""
    t = np.round(np.arange(0.0, T_END, DT), 3)
    if kin is None:
        y = 0.004 * (t / T_END) + 0.0
    else:
        kon, kdis = kin
        kobs = kon * conc_M + kdis
        req = RMAX * kon * conc_M / kobs
        y = np.where(t <= T_ASSOC, req * (1 - np.exp(-kobs * t)), 0.0)
        r_end = req * (1 - math.exp(-kobs * T_ASSOC))
        y = np.where(t > T_ASSOC, r_end * np.exp(-kdis * (t - T_ASSOC)), y)
    return t, y + rng.normal(0.0, NOISE, t.size)


def model(t, kon, kdis, rmax, conc_M):
    kobs = kon * conc_M + kdis
    req = rmax * kon * conc_M / kobs
    r_end = req * (1 - np.exp(-kobs * T_ASSOC))
    return np.where(t <= T_ASSOC, req * (1 - np.exp(-kobs * t)), r_end * np.exp(-kdis * (t - T_ASSOC)))


def fit_row(t, y, conc_M):
    """Local 1:1 fit of one trace -> (kon, kdis, rmax, chi2, r2) or None when nothing binds."""
    if y.max() < 0.03:
        return None
    try:
        popt, _ = curve_fit(lambda tt, kon, kdis, rm: model(tt, kon, kdis, rm, conc_M), t, y,
                            p0=[1e5, 1e-3, y.max()], bounds=([1e2, 1e-7, 0.01], [1e8, 1.0, 5.0]), maxfev=20000)
    except (RuntimeError, ValueError):
        return None
    res = y - model(t, *popt, conc_M)
    sse = float(np.sum(res ** 2))
    return (*popt, sse, 1 - sse / float(np.sum((y - y.mean()) ** 2)))


def sci(x):
    return f"{x:.2E}"


def write_raw(path, conc_M, traces):
    """Octet-style export: header lines, then Time / Data / Sim triplets per analyte, tab-delimited."""
    n = len(traces)
    t = traces[0][0]
    out = ["Vers 3.43 Data Simulation",
           "Conc1\t" + "\t".join([f"{conc_M:.2E}"] * n) + "\t",
           "Start1\t" + "\t".join(["0"] * n) + "\t",
           f"Stop1\t" + "\t".join([f"{T_END - DT:g}"] * n) + "\t",
           "\t".join(f"Time{i + 1}\tData{i + 1}\tSim{i + 1}" for i in range(n)) + "\t"]
    for k in range(t.size):
        cells = []
        for (_, y) in traces:
            cells += [f"{t[k]:g}", f"{y[k]:.8f}", ""]
        out.append("\t".join(cells[:-1]) + "\t")
    with open(path, "w", encoding="latin1", newline="") as fh:
        fh.write("\r\n".join(out) + "\r\n")


HEADER = ["Include", "Sensor Location", "Sensor Type", "Sample ID", "Loading Sample ID", "Conc. (nM)", "Response",
          "X=0", "X=295", "KD (M)", "KD2", "Full  X^2", "Full  R^2", "kon(1/Ms)", "kdis(1/s)", "Rmax"]


def write_data(folder):
    st = S.read(os.path.join(folder, "bli2prism_setup.xlsx"))
    ligands, analytes = st["ligands"], [a["csv_id"] for a in st["analytes"]]
    concs = st["series_M"]
    if len(ligands) != 2 or len(analytes) != 3 or len(concs) != 8:
        raise SystemExit("the demo generator expects 2 ligands, 3 analytes and 8 dilution steps in the setup workbook")
    rng = np.random.default_rng(SEED)
    rows = []
    for col, lig in enumerate(ligands, 1):
        for step, c in enumerate(concs):
            traces = []
            for ai, an in enumerate(analytes, 1):
                t, y = simulate(rng, KINETICS[(col, ai)], c)
                traces.append((t, y))
                loc = f"{'ABCDEFGH'[step]}{col}"
                f = fit_row(t, y, c)
                i295 = int(np.argmin(np.abs(t - 295.0)))
                x295, x0 = float(y[i295]), float(y[0])
                if f is None:                                     # nothing binds: censored values, as Octet writes them
                    rows.append(["x", loc, "Demo sensor", an, lig, f"{c * 1e9:.6g}", f"{x295:.4f}", f"{x0:.5f}",
                                 f"{x295:.5f}", "<1.0E-12", "<1.0E-12", "9.99", "0.0", "<1.0E-12", "<1.0E-07", ""])
                else:
                    kon, kdis, rm, chi, r2 = f
                    rows.append(["x", loc, "Demo sensor", an, lig, f"{c * 1e9:.6g}", f"{x295:.4f}", f"{x0:.5f}",
                                 f"{x295:.5f}", sci(kdis / kon), "<1.0E-12", f"{chi:.6f}", f"{r2:.6f}", sci(kon),
                                 sci(kdis), f"{rm:.4f}"])
            write_raw(os.path.join(folder, f"{'ABCDEFGH'[step]}{col}.xls"), c, traces)
    with open(os.path.join(folder, "kineticanalysistableresults.csv"), "w", newline="") as fh:
        w = csv.writer(fh, lineterminator="\r\n")
        w.writerow(HEADER)
        w.writerows(rows)


def make(folder):
    if not os.path.exists(os.path.join(folder, "bli2prism_setup.xlsx")):
        raise SystemExit(f"{folder} has no bli2prism_setup.xlsx: the generator reads the plate map and dilution series from it")
    write_data(folder)
    return folder


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, "demo_binder_screen")
    print("wrote", make(target))
