"""`bli2prism draft`: write bli2prism_setup.xlsx for a run folder.

Takes the lab's own workbook (its Setup sheet is kept exactly as is, formulas included),
drops the analysis tabs, and adds a pre-filled `bli2prism` sheet. Roles, display names and
'made of' proteins are inferred from the results CSV and the Setup sheet and are meant to be
checked; the two inputs that must not be guessed (kinetic concentrations per ligand, and
confirming the normalization reference) are left for the user.
"""

from __future__ import annotations

import glob
import os

import openpyxl

from .prismwrite import describe_project

from . import refdata as R
from . import setupsheet as S


def draft(folder, source=None, out=None, auto=True):
    out = out or os.path.join(folder, "bli2prism_setup.xlsx")
    if os.path.exists(out):
        raise SystemExit(f"{out} already exists; move or delete it first (draft never overwrites "
                         "your edits)")
    if source is None:
        hits = [f for f in glob.glob(os.path.join(folder, "*.xlsx"))
                if not os.path.basename(f).startswith(("~$", "bli2prism_setup"))]
        if len(hits) != 1:
            raise SystemExit(f"expected exactly one source .xlsx in {folder}, found {hits}; "
                             "pass --source")
        source = hits[0]
    hits = sorted(glob.glob(os.path.join(folder, "kineticanalysistableresults*.csv")))
    if len(hits) != 1:
        raise SystemExit(f"expected exactly one kineticanalysistableresults*.csv in {folder}, found "
                         f"{[os.path.basename(h) for h in hits]}")
    csv_path = hits[0]
    rows = R.read_results_csv(csv_path)
    spec = S.infer(source, rows, os.listdir(folder), auto=auto)
    spec["name"] = os.path.splitext(os.path.basename(source))[0]
    # tell the folder's .prism files apart by what they contain: a populated one is a template
    # (graphs kept), an empty one (as Prism creates it) is the container to fill in
    populated, empty = [], []
    for f in sorted(os.listdir(folder)):
        if f.lower().endswith(".prism"):
            try:
                (empty if describe_project(os.path.join(folder, f))["is_blank"] else populated).append(f)
            except ValueError:
                pass
    spec["prism_template"] = populated[0] if len(populated) == 1 else None
    spec["prism_output"] = empty[0] if len(empty) == 1 else None

    wb = openpyxl.load_workbook(source)
    if S.SETUP not in wb.sheetnames:
        raise SystemExit(f"{source} has no '{S.SETUP}' sheet")
    for name in list(wb.sheetnames):
        if name not in (S.SETUP, S.PROTEINS):      # a Proteins tab (e.g. from an earlier setup) is kept
            del wb[name]
    S.split_proteins(wb)
    S.style_setup(wb[S.SETUP], auto_rows=None)
    S.style_proteins(wb[S.PROTEINS])
    ws_b = S.add_sheet(wb, spec)
    S.style_setup(wb[S.SETUP], auto_rows=S.auto_row_count(ws_b) if auto else None)
    S.link_proteins(wb)
    wb.move_sheet(S.PROTEINS, offset=0)
    S.write_with_cache(out, wb)

    print(f"source  : {os.path.basename(source)}  (Setup sheet kept, other tabs dropped)")
    print(f"written : {out}")
    print("inferred (check the yellow cells on the bli2prism sheet):")
    for k, e in sorted(spec["plate"].items()):
        extra = f"  [{e['check']}]" if e.get("check") else ""
        print(f"  plate column {e['col']:2d}: {e.get('role', '-'):12s} {e.get('display', '') or ''}{extra}")
    print(f"  normalization reference (suggested, largest MW): {spec['norm_ref']}")
    print(f"  Prism template (a populated .prism, graphs kept): {spec['prism_template'] or '(none found)'}")
    print(f"  Prism output file (an EMPTY .prism you uploaded): {spec['prism_output'] or '(none found: enter the name you will upload)'}")
    print(f"you must enter: 'Kinetic concentrations per ligand' and 'Number of analytes' "
          f"(left blank on purpose; the plate map has {spec['analyte_hint']} Analyte column(s))")
