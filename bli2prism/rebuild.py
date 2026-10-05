"""Build a Prism project from raw Octet exports by cloning and reshaping a template.

`build`             any shape the template can hold (see shape.py)
`rebuild-reference` the reference experiment, diffed against the original (expected: no
                    differences)
"""

from __future__ import annotations

import glob
import os

from . import refdata as R
from . import prismwrite as W
from .shape import Project, apply_shape, ShapeError


def find_inputs(folder, setup=None):
    """-> (setup workbook, results csv). Prefers bli2prism_setup.xlsx, then the only .xlsx."""
    from . import setupsheet
    if setup is None:
        own = os.path.join(folder, "bli2prism_setup.xlsx")
        if os.path.exists(own):
            setup = own
        else:
            hits = [f for f in glob.glob(os.path.join(folder, "*.xlsx"))
                    if not os.path.basename(f).startswith("~$")]
            if len(hits) != 1:
                raise FileNotFoundError(f"expected exactly one .xlsx in {folder} (or a "
                                        f"bli2prism_setup.xlsx), found {hits}")
            setup = hits[0]
    name = "kineticanalysistableresults.csv"
    if setupsheet.has_sheet(setup):
        # only the file name is needed here; a sheet that is not ready is reported by whoever
        # reads it properly (build, qc), not by a crash while locating the CSV
        name = setupsheet.read_value(setup, "results_csv") or name
    path = os.path.join(folder, name)
    if not os.path.exists(path):
        have = sorted(os.path.basename(f) for f in glob.glob(os.path.join(folder, "*.csv")))
        raise FileNotFoundError(
            f"the results CSV named on the bli2prism sheet, {name!r}, is not in {folder}"
            + (f" (CSV files there: {have})" if have else " (no CSV files there)")
            + ". Rename the file or correct the 'Results CSV' cell.")
    return setup, path


def build_tables(folder, setup_xlsx, results_csv, n_concentrations=None, snap=True):
    """n_concentrations: how many of each ligand's highest concentrations get a kinetics sheet.

    Deliberately has no default: the number is a manual input per experiment. A ligand that has
    fewer concentrations than that is built with the ones it has, and a NOTE says so (you cannot
    build what the results table does not contain). The equilibrium table always uses every
    concentration the ligand has.
    Returns (tables by sheet title, ligand names, concentration labels).
    """
    rows = R.read_results_csv(results_csv)
    setup = R.read_setup(setup_xlsx, {r["Loading Sample ID"] for r in rows})
    if n_concentrations is None:
        n_concentrations = setup.get("n_concentrations")
    if n_concentrations is None:
        raise ValueError("the number of kinetic concentrations is a manual input: pass -n or fill "
                         "'Kinetic concentrations per ligand' on the bli2prism sheet")
    if n_concentrations < 1:
        raise ValueError("the number of kinetic concentrations must be at least 1")
    series = setup["series_M"]
    an = setup["analytes"]
    files_by_ligand = R.ligand_files(rows, setup["ligands"], series)
    tables, labels = {}, None
    for lig in setup["ligands"]:
        files = files_by_ligand[lig]
        present = sorted(files)
        if not present:
            raise ValueError(f"{lig} has no rows in the results table")
        xs, warns = R.step_concentrations(folder, files, series, snap)
        for w in warns:
            print("WARNING:", w)
        tables[lig] = R.equilibrium_table(rows, lig, series, [xs[s_] for s_ in present], an, steps=present)
        chosen = present[:n_concentrations]
        if len(chosen) < n_concentrations:
            print(f"NOTE: {lig} has only {len(chosen)} concentration(s) in the results table "
                  f"({R.describe_files(files)}); 'Kinetic concentrations per ligand' is {n_concentrations}, "
                  f"so {len(chosen)} are used")
        these = []
        for step in chosen:
            name = files[step]
            raw = R.parse_raw_xls(os.path.join(folder, name))
            hdr = raw["concs"]
            if max(hdr) / min(hdr) > 1.0001:
                raise ValueError(f"{name}: traces disagree on Conc1 {hdr}")
            got = R.match_step(hdr[0], series)
            if got != step:
                raise ValueError(f"{name} has Conc1 {hdr[0]:g} M, which is Setup dilution step {got + 1}, but the "
                                 f"results table puts it at step {step + 1}")
            label = R.conc_label_nM(hdr[0])
            these.append(label)
            tables[f"Kinetics {lig} - {label}"] = R.kinetics_table(raw, an)
        if labels is not None and these != labels:
            raise ValueError(f"{lig}: concentration labels {these} differ from {labels}")
        labels = these
    return tables, setup["ligands"], labels


def find_one(folder, pattern):
    hits = glob.glob(os.path.join(folder, pattern))
    if len(hits) != 1:
        hint = ""
        if pattern == "*.prism":
            hint = (". Name the template with -t <file>, or build without one using --no-template "
                    "(add --blank <file> to use your own blank project as the container)")
        raise FileNotFoundError(f"expected exactly one {pattern} in {folder}, found {hits}{hint}")
    return hits[0]


def find_template(folder):
    """The folder's one .prism, or None when it has none (then the project is built from
    scratch). Several is ambiguous and an error: guessing would silently drop the graphs."""
    hits = glob.glob(os.path.join(folder, "*.prism"))
    if not hits:
        return None
    return find_one(folder, "*.prism")            # raises, with a hint, if there are several


def _prism_name(name, what):
    name = str(name).strip()
    if os.sep in name or "/" in name or name in ("", ".", ".."):
        raise ValueError(f"'{what}' must be a file name, not a path: {name!r}")
    return name if name.lower().endswith(".prism") else name + ".prism"


def resolve_prism(folder, setup_path=None, template=None, no_template=False, blank=None):
    """Decide what to build from. Returns {template, container, out_name, why}.

    Precedence (the sheet's cells are the default, command-line options override):
      1. a template (-t, or the sheet's 'Prism template'): the template route, graphs kept
      2. otherwise an empty project named by the sheet's 'Prism output file' (or --blank) that is
         present in the folder: the no-template route with that project as the container
      3. otherwise the folder's only .prism (template route), or none (built from scratch with
         the bundled container). Several .prism files is an error: guessing would silently
         drop the graphs.
    --no-template skips 1 and 3.
    """
    from . import setupsheet
    sheet_tmpl = sheet_out = None
    if setup_path and setupsheet.has_sheet(setup_path):
        sheet_tmpl = setupsheet.read_value(setup_path, "prism_template")
        sheet_out = setupsheet.read_value(setup_path, "prism_output")
    out_name = _prism_name(sheet_out, "Prism output file") if sheet_out else None
    res = {"template": None, "container": None, "out_name": out_name, "why": ""}
    if not no_template:
        want = template or (os.path.join(folder, _prism_name(sheet_tmpl, "Prism template")) if sheet_tmpl else None)
        if want:
            if not os.path.exists(want):
                src = "-t" if template else "the 'Prism template' cell"
                raise FileNotFoundError(f"the template named by {src}, {os.path.basename(want)!r}, is not in "
                                        f"{folder}. Upload it, or clear the cell to build without a template")
            res.update(template=want, why="template named " + ("on the command line" if template else "on the setup sheet"))
            return res
    named = blank or (os.path.join(folder, out_name) if out_name else None)
    if named and os.path.exists(named):
        res.update(container=named, why="empty project " + ("given with --blank" if blank else f"named {out_name!r} on the setup sheet"))
        return res
    if blank:
        raise FileNotFoundError(f"--blank {blank!r} does not exist")
    if no_template:
        res["why"] = "no template, bundled container"
        return res
    found = find_template(folder)
    res.update(template=found, why="the folder's only .prism" if found else "no .prism in the folder: built from scratch")
    return res


def build(folder, template, out, n_concentrations, setup=None, scratch=False, container=None):
    """template=None (or scratch=True) builds without a template: see scratch.py."""
    setup_xlsx, results_csv = find_inputs(folder, setup)
    tables, ligands, labels = build_tables(folder, setup_xlsx, results_csv, n_concentrations)
    analytes = [a["display"] for a in R.read_setup(setup_xlsx, {r["Loading Sample ID"] for r in R.read_results_csv(results_csv)})["analytes"]]
    if scratch or template is None:
        from .scratch import build_scratch
        proj, _ = build_scratch(tables, ligands, labels, container, analytes)
    else:
        proj = Project(template)
        apply_shape(proj, ligands, labels)
        idx = {sj["title"]: u for u, sj in proj.data_sheets().items()
               if u in proj.doc["sheets"]["data"]}
        for title, lines in tables.items():
            if title not in idx:
                raise ShapeError(f"no sheet titled {title!r} after reshaping")
            proj.put_table(idx[title], lines)
    problems = proj.validate()
    if out:                      # out=None is a dry run (used by qc)
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        proj.save(out)
    return proj, problems, ligands, labels


def report(template, out, proj, problems, ligands, labels):
    print("template : " + (os.path.basename(template) if template else "(none: built from scratch)"))
    print(f"written  : {out}")
    print(f"shape    : {len(ligands)} ligand(s) x {len(labels)} kinetic concentration(s) "
          f"{labels}")
    for w in proj.warnings:
        print("NOTE:", w)
    if problems:
        print(f"VALIDATION PROBLEMS ({len(problems)}):")
        for p in problems:
            print("  -", p)
    else:
        print("validation: no structural problems found")


def unused_path(path):
    """`path`, or `name_2.ext`, `name_3.ext`, ... if it exists: a build never overwrites a file, because the one
    already there may be a previous build the user has since opened in Prism and added graphs to."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 2
    while os.path.exists(f"{base}_{n}{ext}"):
        n += 1
    return f"{base}_{n}{ext}"


def build_command(folder, out, n_concentrations, template=None, setup=None, no_template=False,
                  blank=None):
    setup_path, _ = find_inputs(folder, setup)
    pick = resolve_prism(folder, setup_path, template, no_template, blank)
    out = out or os.path.join(folder, "rebuild", pick["out_name"] or "build.prism")
    wanted, out = out, unused_path(out)
    if out != wanted:
        print(f"NOTE: {os.path.basename(wanted)} already exists and was left untouched; writing {os.path.basename(out)}")
    proj, problems, ligands, labels = build(folder, pick["template"], out, n_concentrations, setup,
                                            scratch=pick["template"] is None, container=pick["container"])
    print(f"source   : {pick['why']}" + (f" ({os.path.basename(pick['container'])})" if pick["container"] else ""))
    report(pick["template"], out, proj, problems, ligands, labels)
    return problems
