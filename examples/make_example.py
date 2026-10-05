"""Generate the demo experiment in examples/demo_binder_screen/ from nothing but random numbers.

No real data is involved: the traces are simulated 1:1 binding curves with seeded noise, the proteins and
their molecular weights are invented, and the results table is produced by fitting the simulated traces.

    python3 examples/make_example.py                 # rewrites examples/demo_binder_screen/
    python3 examples/make_example.py OUT_FOLDER      # somewhere else

What it writes (the same files a real Octet run folder holds):
    A1.xls .. H1.xls      sensor column 1 = Binder A, plate rows A-H = the 8 analyte concentrations
    A2.xls .. H2.xls      sensor column 2 = Binder B
    kineticanalysistableresults.csv   one fit row per ligand / analyte / concentration
    bli2prism_setup.xlsx  the filled-in setup workbook (bli2prism, Setup and Proteins tabs)
"""

from __future__ import annotations

import csv
import math
import os
import shutil
import sys
import tempfile

import numpy as np
import openpyxl
from openpyxl.styles import Alignment
from scipy.optimize import curve_fit

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from bli2prism import setupsheet as S  # noqa: E402
from bli2prism.draft import draft  # noqa: E402

SEED = 20260101
LIGANDS = ["Binder A", "Binder B"]                       # sensor columns 1 and 2
ANALYTES = ["Target X", "Target X/Partner Z", "Partner Z"]
DISPLAY = {"Target X": "Target X", "Target X/Partner Z": "Target X/Partner Z complex", "Partner Z": "Partner Z"}
PROTEINS = {                                              # name: (A280, MW g/mol, extinction coefficient)
    "Binder A": (1.50, 26000.0, 31000), "Binder B": (1.20, 24500.0, 29000), "Control IgG": (1.40, 150000.0, 210000),
    "Target X": (0.40, 18000.0, 15000), "Partner Z": (0.55, 27000.0, 22000),
}
START_M, FACTOR, N_STEPS = 1e-6, 3.0, 8
T_ASSOC, T_END, DT = 300.0, 600.0, 0.5
RMAX = 0.55                                               # nm
# (kon 1/Ms, kdis 1/s) per ligand and analyte; None = no binding (only drift and noise)
KINETICS = {
    ("Binder A", "Target X"): (2.0e5, 1.0e-3),             # KD  5 nM
    ("Binder A", "Target X/Partner Z"): (4.0e5, 4.0e-4),   # KD  1 nM
    ("Binder A", "Partner Z"): None,
    ("Binder B", "Target X"): (1.0e5, 5.0e-3),             # KD 50 nM
    ("Binder B", "Target X/Partner Z"): (2.0e5, 4.0e-3),   # KD 20 nM
    ("Binder B", "Partner Z"): (2.0e4, 4.0e-2),            # KD  2 uM (weak)
}
NOISE = 0.002


def concentrations():
    return [START_M / FACTOR ** i for i in range(N_STEPS)]


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
    rng = np.random.default_rng(SEED)
    concs = concentrations()
    rows = []
    for col, lig in enumerate(LIGANDS, 1):
        for step, c in enumerate(concs):
            traces = []
            for an in ANALYTES:
                t, y = simulate(rng, KINETICS[(lig, an)], c)
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


def write_source(path):
    """A lab-style workbook: a Setup tab (plate map, dilution series) and a Proteins tab, nothing else."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Setup"
    plate = {2: "Buffer", 3: "Binder A", 4: "Binder B", 5: "Control IgG", 6: "Buffer", 7: "Buffer", 8: "Buffer",
             9: ANALYTES[0], 10: ANALYTES[1], 11: ANALYTES[2], 12: "10 mM Glycine pH 1.5", 13: "Buffer"}
    for c, text in plate.items():
        ws.cell(5, c, text)
        ws.merge_cells(start_row=5, start_column=c, end_row=12, end_column=c)
        ws.cell(5, c).alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    for r, letter in enumerate("ABCDEFGH", 5):
        ws.cell(r, 1, letter)
    for c in range(2, 14):
        ws.cell(4, c, c - 1)
    ws["N4"], ws["O4"], ws["P4"] = "[concentration] (M)", "[concentration] (nM)", "log"
    ws["N5"] = START_M
    ws["N6"] = START_M
    for r in range(7, 13):
        ws[f"N{r}"] = f"=N{r - 1}/$R$5"
    for r in range(5, 13):
        ws[f"O{r}"] = f"=N{r}*10^9"
        ws[f"P{r}"] = f"=LOG(N{r})"
    ws["Q4"], ws["Q5"], ws["R5"] = "Dilution settings", "Dilution Factor", FACTOR
    ws["Q6"], ws["R6"] = "Transfer volume", "=R7/(R5-1)"
    ws["Q7"], ws["R7"] = "Final Volume per Well", 190
    ws["Q8"], ws["R8"] = "Final Volume + Transfer Volume", "=SUM(R6:R7)"
    pr = wb.create_sheet("Proteins")
    heads = ["Protein", "A280", "Molecular weight (g/mol)", "Extinction Coefficient", "Concentration (M)",
             "Concentration (mg/mL)", "Moles Needed", "Desired Concentration (mol/L)", "Total Volume (uL)",
             "Vol Protein (uL)", "Vol buffer (ul)", "Vol buffer (ul) / 2", "Mass protein needed (ug)"]
    for c, h in enumerate(heads, 1):
        pr.cell(4, c, h)
    for r, (name, (a280, mw, ext)) in enumerate(PROTEINS.items(), 5):
        pr.cell(r, 1, name), pr.cell(r, 2, a280), pr.cell(r, 3, mw), pr.cell(r, 4, ext)
        pr.cell(r, 5, f"=B{r}/D{r}"), pr.cell(r, 6, f"=E{r}*C{r}")
        pr.cell(r, 10, f"=G{r}/E{r}*10^6"), pr.cell(r, 11, f"=I{r}-J{r}"), pr.cell(r, 12, f"=K{r}/2")
        pr.cell(r, 13, f"=F{r}*J{r}")
    wb.save(path)


def write_setup(folder):
    """Draft the setup workbook from the source, then fill in the inputs that must be entered by hand."""
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "source.xlsx")
        write_source(src)
        shutil.copy(os.path.join(folder, "kineticanalysistableresults.csv"), tmp)
        out = os.path.join(folder, "bli2prism_setup.xlsx")
        import contextlib
        import io
        with contextlib.redirect_stdout(io.StringIO()):
            draft(tmp, source=src, out=out)
    wb = openpyxl.load_workbook(out)
    ws = wb[S.SHEET]
    labels = {ws[f"A{r}"].value: r for r in range(5, 16) if ws[f"A{r}"].value}
    p0 = S.plate_row0()
    for r in range(p0, p0 + 40):                              # display names for the analytes
        text = ws[f"D{r}"].value
        if text in DISPLAY and text != DISPLAY[text]:
            ws[f"D{r}"] = DISPLAY[text]
    for label, value in (("Experiment name", "Demo binder screen"), ("Results CSV", "kineticanalysistableresults.csv"),
                         ("Kinetic concentrations per ligand", N_STEPS), ("Number of analytes", len(ANALYTES)),
                         ("Normalization reference analyte", ANALYTES[1]), ("Dilution steps", N_STEPS)):
        ws[f"B{labels[label]}"] = value
    S.write_with_cache(out, wb)
    return out


def make(out_folder):
    os.makedirs(out_folder, exist_ok=True)
    write_data(out_folder)
    return write_setup(out_folder)


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, "demo_binder_screen")
    print("wrote", make(target))
