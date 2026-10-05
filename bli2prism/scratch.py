"""Build a Prism project with no user-supplied template.

The data sheets and the fit analyses are plain JSON + CSV, so they are cloned, with fresh UIDs,
fresh fenIDs and the new titles, from the skeleton project bundled with the package,
bli2prism/_resources/prism_skeleton.prism (built from simulated data, no graphs or layouts). The shape is unlimited: any number of ligands and concentrations.

What is NOT written: graphs and layouts. Their binary files cannot be generated, so you
create the graphs in Prism (they are linked to nothing here, so there is nothing to break).

The container (document.json, info sheet, fonts) comes from the bundled skeleton, or from a
blank project you saved in Prism, in which case your Prism version's own skeleton is kept.
"""

from __future__ import annotations

import copy
import json
import os
import re
import uuid

from .shape import Project, ShapeError, UID_RE

BLUEPRINT = os.path.join(os.path.dirname(__file__), "_resources", "prism_skeleton.prism")
LOWER_UUID = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}")
NS = uuid.UUID("6f1c2f5e-9d4b-4c52-8f3e-0b1a7d3c9e11")

# skeleton objects that get cloned
BP_LIGAND = "Binder A"
BP_KINETICS = "Kinetics Binder A - 1000 nM"
MAX_ANALYTES = 7
GREEK = re.compile(r"[\u0370-\u03ff]")


def _new_uid():
    return str(uuid.uuid4()).upper()


# ---------------------------------------------------------------- container handling
def wipe_sheets(p):
    """Remove every data sheet, graph, layout and analysis, keeping the document skeleton."""
    for a in list(p.doc["sheets"].get("analyses", [])):
        p.remove_analysis(a)
    for g in list(p.doc["sheets"].get("graphs", [])):
        p.remove_graph(g)
    for l in list(p.doc["sheets"].get("layouts", [])):
        p._remove_files([n for n in p.files if n.startswith(f"layouts/{l}/")])
        p._forget(l)
    for d in list(p.doc["sheets"].get("data", [])):
        p.remove_data_sheet(d)
    # drop everything left under data/ (sets, result tables) so nothing orphaned survives
    keep = {"data/", "data/sheets/", "data/tables/", "data/sets/"}
    for n in [n for n in p.files if n.startswith("data/") and n not in keep]:
        p.files.pop(n)
        p.infos.pop(n, None)
        p.style.pop(n, None)
    p.doc["sheetAttributesMap"] = {k: v for k, v in p.doc["sheetAttributesMap"].items()
                                   if k in p.doc["sheets"].get("info", [])}


def _ensure_dirs(p, name, info_like):
    parts = name.split("/")[:-1]
    for i in range(1, len(parts) + 1):
        d = "/".join(parts[:i]) + "/"
        if d not in p.files:
            p.files[d] = b""
            zi = copy.copy(info_like)
            zi.filename = d
            p.infos[d] = zi


def _max_fenids(p):
    out = {}
    for n, data in p.files.items():
        if not n.endswith(".json"):
            continue
        def walk(o):
            if isinstance(o, dict):
                if isinstance(o.get("fenID"), int):
                    k = o.get("@class")
                    out[k] = max(out.get(k, -1), o["fenID"])
                for v in o.values():
                    walk(v)
            elif isinstance(o, list):
                for v in o:
                    walk(v)
        walk(json.loads(data))
    return out


def clone_group(src, dst, names, replace=()):
    """Copy the member files `names` from project src into dst with fresh UIDs / fenIDs.

    replace: (old, new) pairs applied to every string value of the copied JSON (titles).
    Returns {old uid: new uid}. Every UID found in the group is remapped, so the group must be
    closed: nothing in it may refer to an object outside it.
    """
    files = [n for n in names if not n.endswith("/")]
    uids = set()
    for n in files:
        uids |= set(UID_RE.findall(n))
        if n.endswith(".json"):
            uids |= set(UID_RE.findall(src.files[n].decode("utf-8")))
    umap = {u: _new_uid() for u in sorted(uids)}
    counters = _max_fenids(dst)
    out_names = []
    for n in files:
        new_name = n
        for o, w in umap.items():
            new_name = new_name.replace(o, w)
        data = src.files[n]
        if n.endswith(".json"):
            text = data.decode("utf-8")
            for o, w in umap.items():
                text = text.replace(o, w)
            for o in set(LOWER_UUID.findall(text)):          # column / model ids
                text = text.replace(o, str(uuid.uuid5(NS, new_name + o)))
            src.jget(n)                                   # populates src.style[n]
            obj = json.loads(text)

            def walk(x):
                if isinstance(x, dict):
                    if isinstance(x.get("fenID"), int):
                        k = x.get("@class")
                        counters[k] = counters.get(k, -1) + 1
                        x["fenID"] = counters[k]
                    return {k: walk(v) for k, v in x.items()}
                if isinstance(x, list):
                    return [walk(v) for v in x]
                if isinstance(x, str):
                    for o, w in replace:
                        if o in x:
                            x = x.replace(o, w)
                return x
            obj = walk(obj)
            style = src.style[n]
            ind, nl = style
            data = (json.dumps(obj, indent=ind, ensure_ascii=False) + nl).encode("utf-8")
            dst.style[new_name] = style
        zi = copy.copy(src.infos[n])
        zi.filename = new_name
        _ensure_dirs(dst, new_name, zi)
        dst.files[new_name] = data
        dst.infos[new_name] = zi
        out_names.append(new_name)
    return umap


def _title_of_sheet(p, uid):
    return p.jget(f"data/sheets/{uid}/sheet.json")["title"]


# ----------------------------------------------------------------------- analyte count
def _is_greek(ch):
    return bool(GREEK.match(ch))


def make_title(proto, text):
    """A dataset title object for `text`, formatted like prototype title `proto`: Greek letters
    get the Symbol font, as Prism stores them; everything else uses the prototype's base run."""
    if isinstance(proto, str):
        return text
    runs = proto.get("formatting") or [{"attributes": {}}]
    base = {k: v for k, v in runs[0]["attributes"].items() if k != "fontFamily"}
    out, i = [], 0
    while i < len(text):
        j = i
        while j + 1 < len(text) and _is_greek(text[j + 1]) == _is_greek(text[i]):
            j += 1
        attrs = dict(base)
        if _is_greek(text[i]):
            attrs = {"fontFamily": "Symbol", **base}
        out.append({"range": f"{i}~{j}" if j > i else f"{i}", "attributes": attrs})
        i = j + 1
    return {"string": text, "version": proto.get("version", 1), "formatting": out}


def _title_string(t):
    return t if isinstance(t, str) else t["string"]


def resize_y_sets(dst, sheet_uid, titles):
    """Make a data sheet have exactly len(titles) y datasets named `titles`.

    Surplus datasets are dropped; extra ones are cloned from the last. Returns the sheet's
    dataset uid list (also written back to the sheet)."""
    name = f"data/sheets/{sheet_uid}/sheet.json"
    sj = dst.jget(name)
    t = sj["table"]
    ys = list(t["dataSets"])
    n = len(titles)
    for u in ys[n:]:
        dst._remove_files([f"data/sets/{u}.json"])
    ys = ys[:n]
    while len(ys) < n:
        proto = ys[-1]
        pf = f"data/sets/{proto}.json"
        if pf in dst.files:
            umap = clone_group(dst, dst, [pf])
            ys.append(umap[proto])
        else:
            ys.append(_new_uid())
    for u, title in zip(ys, titles):
        f = f"data/sets/{u}.json"
        if f in dst.files:
            d = dst.jget(f)
            if _title_string(d["title"]) != title:
                d["title"] = make_title(d["title"], title)
                dst.jput(f, d)
    t["dataSets"] = ys
    dst.jput(name, sj)
    return ys


def _set_columns(dst, sheet_uid, n_cols):
    sj = dst.jget(f"data/sheets/{sheet_uid}/sheet.json")
    cj_name = f"data/tables/{sj['table']['uid']}/content.json"
    cj = dst.jget(cj_name)
    if cj["numberOfColumns"] != n_cols:
        cj["numberOfColumns"] = n_cols
        dst.jput(cj_name, cj)


def _letters(j):
    return chr(ord("A") + j)


def _col(uid, title, j):
    return {"idx": j, "title": title, "uid": uid}


def resize_data_sheet(dst, sheet_uid, titles):
    """Equilibrium or kinetics sheet: x column + one column per analyte."""
    resize_y_sets(dst, sheet_uid, titles)
    _set_columns(dst, sheet_uid, 1 + len(titles))


def resize_analysis(dst, analysis_uid, eq_sets, titles):
    """Re-point an analysis at `eq_sets` (one per analyte, in order) and resize everything in
    it that has one entry per analyte: its input list, initial values, results models, and
    its three result tables."""
    n = len(titles)
    a_name = f"analyses/{analysis_uid}/sheet.json"
    a = dst.jget(a_name)
    proto = a["inputDataSets"]
    a["inputDataSets"] = [{"uid": eq_sets[j], "title": titles[j], "idx": _letters(j)} for j in range(n)]
    dst.jput(a_name, a)

    # initial values: Bottom / Top / LogEC50, each a list with one entry per analyte
    p_name = f"analyses/{analysis_uid}/parameters.json"
    par = dst.jget(p_name)
    iv = par["content"]["model"]["initialValues"]
    for key, lst in list(iv.items()):
        new = []
        for j in range(n):
            e = copy.deepcopy(lst[min(j, len(lst) - 1)])
            e["dataSet"]["col"] = _col(eq_sets[j], titles[j], j)
            new.append(e)
        iv[key] = new
    dst.jput(p_name, par)

    # results: one model and one data summary per analyte, keyed by a per-analyte uuid
    r_name = f"analyses/{analysis_uid}/results.json"
    res = dst.jget(r_name)
    content = res["content"]

    def rekey(d, tag):
        old = list(d.values())
        new = {}
        for j in range(n):
            e = copy.deepcopy(old[min(j, len(old) - 1)])
            e["id"]["col"] = _col(eq_sets[j], titles[j], j)
            txt = json.dumps(e)
            txt = re.sub(r'"dsList": \[\s*\d+\s*\]', f'"dsList": [{j}]', txt)
            new[str(uuid.uuid5(NS, f"{analysis_uid}/{tag}/{j}"))] = json.loads(txt)
        return new
    content["models"] = rekey(content["models"], "model")
    content["dataSummary"] = rekey(content["dataSummary"], "summary")
    dst.jput(r_name, res)

    # result tables
    for rn in sorted(x for x in dst.files if x.startswith(f"analyses/{analysis_uid}/result_sheets/")
                     and x.endswith(".json")):
        ds = dst.jget(rn)["dataSheet"]
        sj = dst.jget(f"data/sheets/{ds}/sheet.json")
        t = sj["table"]
        resize_y_sets(dst, ds, titles)
        csv_name = f"data/tables/{t['uid']}/data.csv"
        text = dst.files[csv_name].decode("utf-8")
        if t.get("dataFormat") == "text":                       # table of results: row titles + n columns
            _set_columns(dst, ds, 1 + n)
            dst.files[csv_name] = _reshape_csv(text, 1 + n).encode("utf-8")
        elif t.get("format") == "xy":                           # fitted curves: n columns, x is a series
            _set_columns(dst, ds, n)
            dst.files[csv_name] = _reshape_csv(text, n).encode("utf-8")
            sj = dst.jget(f"data/sheets/{ds}/sheet.json")
            vs = sj["table"].get("valuesStorageData")
            if vs:
                vs["last rows"] = [f"{j};999" for j in reversed(range(n))]
                vs["yColumnUuid"] = [str(uuid.uuid5(NS, f"{ds}/y/{j}")) for j in range(n)]
                dst.jput(f"data/sheets/{ds}/sheet.json", sj)
        else:                                                   # internal curve view: empty
            _set_columns(dst, ds, 1 + n)


def _reshape_csv(text, n_cols):
    """Truncate or widen every line to n_cols fields, repeating the last field."""
    out = []
    for ln in text.split("\n"):
        if ln == "":
            out.append(ln)
            continue
        f = ln.split(",")
        f = f[:n_cols] + [f[-1]] * max(0, n_cols - len(f))
        out.append(",".join(f))
    return "\n".join(out)


# ------------------------------------------------------------------------- the build
def equilibrium_title(ligand):
    """Title of a ligand's equilibrium data sheet in a from-scratch project."""
    return f"{ligand} Equilibrium"


def build_scratch(tables, ligands, labels, container=None, analytes=None):
    """tables: {sheet title: csv lines}, as from rebuild.build_tables (a ligand's equilibrium table is keyed by the ligand's
    name; the sheet is titled equilibrium_title(ligand)).

    analytes: display names, one per analyte (1 to MAX_ANALYTES); default = the skeleton's 3.
    Returns (project, warnings). Sheets are named exactly as in the template-based build.
    """
    if analytes is not None and not 1 <= len(analytes) <= MAX_ANALYTES:
        raise ShapeError(f"{len(analytes)} analytes requested; 1 to {MAX_ANALYTES} are supported")
    bp = Project(BLUEPRINT)
    dst = Project(container) if container else Project(BLUEPRINT)
    had_keys = set(dst.doc["sheets"])            # keys Prism itself wrote for this container
    if container:
        n_data, n_an = len(dst.doc["sheets"].get("data", [])), len(dst.doc["sheets"].get("analyses", []))
        if n_data > 1 or n_an:
            dst.warnings.append(
                f"the container project is not empty ({n_data} data sheets, {n_an} analyses): its "
                "sheets are not carried over (the output starts from its skeleton only). Your original "
                "file is untouched")
    wipe_sheets(dst)
    block = bp.ligand_blocks()[0]
    if block["ligand"] != BP_LIGAND:
        raise ShapeError("bundled skeleton is not the expected project")
    kin_uid = next(k for k, _ in block["kinetics"] if _title_of_sheet(bp, k) == BP_KINETICS)

    new_eq, new_analyses, new_kin, new_results = [], [], [], []
    for lig in ligands:
        names = list(bp._sheet_files(block["eq_sheet"]))
        for a in block["analysis"]:
            names += [n for n in bp.files if n.startswith(f"analyses/{a}/")]
            for n in [n for n in bp.files if n.startswith(f"analyses/{a}/result_sheets/") and n.endswith(".json")]:
                names += bp._sheet_files(bp.jget(n)["dataSheet"])
        umap = clone_group(bp, dst, names, replace=[(BP_LIGAND, lig)])
        eq_path = f"data/sheets/{umap[block['eq_sheet']]}/sheet.json"
        eq = dst.jget(eq_path)
        eq["title"] = equilibrium_title(lig)
        dst.jput(eq_path, eq)
        new_eq.append(umap[block["eq_sheet"]])
        new_analyses.append(umap[block["analysis"][0]])
        if analytes is not None:
            eq_sets = resize_y_sets(dst, umap[block["eq_sheet"]], analytes)
            _set_columns(dst, umap[block["eq_sheet"]], 1 + len(analytes))
            resize_analysis(dst, umap[block["analysis"][0]], eq_sets, analytes)
    for lig in ligands:
        row = []
        for label in labels:
            title = f"Kinetics {lig} - {label}"
            umap = clone_group(bp, dst, bp._sheet_files(kin_uid), replace=[(BP_KINETICS, title)])
            row.append(umap[kin_uid])
            if analytes is not None:
                resize_data_sheet(dst, umap[kin_uid], analytes)
        new_kin.append(row)

    # register in document.json: eq sheets, then kinetics per ligand; analyses per ligand
    doc = dst.doc
    doc["sheets"]["data"] = new_eq + [k for row in new_kin for k in row]
    doc["sheets"]["analyses"] = new_analyses
    doc["sheets"]["graphs"], doc["sheets"]["layouts"] = [], []
    for key in ("graphs", "layouts"):            # Prism omits empty lists: do the same
        if key not in had_keys:
            doc["sheets"].pop(key)
    for uid in doc["sheets"]["data"]:
        doc["sheetAttributesMap"][uid] = {"title": _title_of_sheet(dst, uid)}
    for uid in new_analyses:
        doc["sheetAttributesMap"][uid] = {"title": dst.jget(f"analyses/{uid}/sheet.json")["title"]}
    n_sheets = len(doc["sheets"]["data"]) + 3 * len(new_analyses)
    doc["uiSettings"]["nextDataSheetNumber"] = n_sheets + 1
    cur = doc["uiSettings"].setdefault("currentSheets", {})
    cur["data"], cur["analysis"] = doc["sheets"]["data"][0], new_analyses[0]
    cur.pop("graph", None)
    cur.pop("layout", None)
    doc["uiSettings"]["currentSheetType"] = "Data"      # there are no layouts or graphs to show

    # point every table at its data
    idx = {_title_of_sheet(dst, u): u for u in doc["sheets"]["data"]}
    for title, lines in tables.items():
        if title in ligands:
            title = equilibrium_title(title)
        if title not in idx:
            raise ShapeError(f"no sheet titled {title!r} was created")
        dst.put_table(idx[title], lines)
    dst.warnings.append("no graphs or layouts were written (they cannot be generated); create the "
                        "graphs in Prism from the data sheets and fit analyses")
    return dst, dst.warnings
