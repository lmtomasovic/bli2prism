"""Change the shape of a cloned .prism project: ligands, concentrations, row counts.

Supported
    - rename ligands and relabel concentrations (sheet titles, graph titles, and the title
      text stored inside each graph's binary)
    - fewer ligands / fewer concentrations than the template (surplus sheets, graphs and
      analyses are removed)
    - different row counts (time points per sensorgram, points per equilibrium curve)

Not supported (raises ShapeError instead of guessing)
    - more ligands or concentrations than the template holds. A graph's binary refers to
      its own fenID and to its datasets' fenIDs in many places, plus cached axis values,
      so a graph cannot be synthesized safely. Build the template with the largest shape
      you will need and let the tool remove the surplus.
    - a different number of analytes (each graph binary holds one series record per analyte).

Layouts hold one placement record per graph in a binary whose object counts and ID tables
are not understood, so a layout is kept only if every graph it links still exists;
otherwise it is removed and reported, to be recreated in Prism.
"""

from __future__ import annotations

import json
import re
import struct
import zipfile

UID_RE = re.compile(r"[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}")


class ShapeError(Exception):
    pass


class Project:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.infos = {i.filename: i for i in z.infolist()}
            self.files = {i.filename: z.read(i) for i in z.infolist()}
        self.style = {}
        self.warnings = []
        self.doc = self.jget("document.json")

    # ------------------------------------------------------------------ json io
    def jget(self, name):
        raw = self.files[name].decode("utf-8")
        obj = json.loads(raw)
        if name not in self.style:
            for ind, nl in (("\t", ""), (4, "\n")):
                if json.dumps(obj, indent=ind, ensure_ascii=False) + nl == raw:
                    self.style[name] = (ind, nl)
                    break
            else:
                raise ShapeError(f"{name}: unrecognised JSON formatting; refusing to rewrite")
        return obj

    def jput(self, name, obj):
        ind, nl = self.style[name]
        new = (json.dumps(obj, indent=ind, ensure_ascii=False) + nl).encode("utf-8")
        if new != self.files[name]:
            self.files[name] = new

    def sync_doc(self):
        self.jput("document.json", self.doc)

    # ------------------------------------------------------------------ indexes
    def data_sheets(self):
        """uid -> sheet.json object, for every data sheet (incl. analysis result sheets)."""
        out = {}
        for n in self.files:
            if n.startswith("data/sheets/") and n.endswith("/sheet.json"):
                d = self.jget(n)
                out[d["uid"]] = d
        return out

    def graph_sheets(self):
        return {self.jget(n)["uid"]: self.jget(n) for n in self.files
                if n.startswith("graphs/") and n.endswith("/sheet.json")}

    def analyses(self):
        return {self.jget(n)["uid"]: self.jget(n) for n in self.files
                if n.startswith("analyses/") and n.count("/") == 2 and n.endswith("/sheet.json")}

    def ligand_blocks(self):
        """Ordered list of dicts describing each ligand's sheets, graphs and analysis."""
        ds, gs, an = self.data_sheets(), self.graph_sheets(), self.analyses()
        title = lambda u: ds[u]["title"]
        data_uids = self.doc["sheets"]["data"]
        graph_by_title = {}
        for g in self.doc["sheets"]["graphs"]:
            graph_by_title[gs[g]["title"]] = g
        blocks = []
        for u in data_uids:
            if title(u).startswith("Kinetics "):
                continue
            lig = title(u)
            kin = [k for k in data_uids if title(k).startswith(f"Kinetics {lig} - ")]
            analysis = [a for a, d in an.items()
                        if any(s["uid"] == u for s in d.get("inputSheets", []))]
            blocks.append({
                "ligand": lig, "eq_sheet": u, "eq_graph": graph_by_title.get(lig),
                "kinetics": [(k, graph_by_title.get(title(k))) for k in kin],
                "analysis": analysis,
            })
        return blocks

    # --------------------------------------------------------------- file level
    def _sheet_files(self, uid):
        sj = self.jget(f"data/sheets/{uid}/sheet.json")
        t = sj["table"]
        sets = list(t.get("dataSets", []))
        for k in ("xDataSet", "rowTitlesDataSet"):
            if t.get(k):
                sets.append(t[k])
        names = [n for n in self.files if n.startswith(f"data/sheets/{uid}/")
                 or n.startswith(f"data/tables/{t['uid']}/")]
        names += [f"data/sets/{s}.json" for s in sets if f"data/sets/{s}.json" in self.files]
        return names

    def _remove_files(self, names):
        for n in names:
            self.files.pop(n, None)
            self.infos.pop(n, None)
            self.style.pop(n, None)
        # drop directory entries that no longer contain anything
        for d in [n for n in self.files if n.endswith("/")]:
            if d in ("data/", "data/sheets/", "data/tables/", "data/sets/", "graphs/",
                     "analyses/", "layouts/", "info/", "misc/"):
                continue
            if not any(m.startswith(d) and m != d for m in self.files):
                self.files.pop(d)
                self.infos.pop(d, None)

    def _forget(self, uid):
        self.doc["sheetAttributesMap"].pop(uid, None)
        for lst in self.doc["sheets"].values():
            if uid in lst:
                lst.remove(uid)

    def remove_data_sheet(self, uid):
        self._remove_files(self._sheet_files(uid))
        self._forget(uid)

    def remove_graph(self, uid):
        self._remove_files([n for n in self.files if n.startswith(f"graphs/{uid}/")])
        self._forget(uid)

    def remove_analysis(self, uid):
        for n in [n for n in self.files if n.startswith(f"analyses/{uid}/result_sheets/")
                  and n.endswith(".json")]:
            ds = self.jget(n)["dataSheet"]
            if f"data/sheets/{ds}/sheet.json" in self.files:
                self.remove_data_sheet(ds)
        self._remove_files([n for n in self.files if n.startswith(f"analyses/{uid}/")])
        self._forget(uid)

    # ------------------------------------------------------------------- titles
    @staticmethod
    def _splice_title(b, old, new, tag, nul):
        """Replace one length-prefixed string record in a graph binary.

        nul=True : tag 0x0009, u32 length includes a trailing NUL (sheet title)
        nul=False: tag 0x002c, u32 length is the character count (chart title)
        """
        o, n = old.encode("latin1"), new.encode("latin1")
        pre = struct.pack("<H", tag)
        pat = pre + struct.pack("<I", len(o) + nul) + o + (b"\x00" if nul else b"")
        if b.count(pat) != 1:
            raise ShapeError(f"graph binary: expected exactly one {old!r} record "
                             f"(tag {tag:#x}), found {b.count(pat)}")
        rep = pre + struct.pack("<I", len(n) + nul) + n + (b"\x00" if nul else b"")
        return b.replace(pat, rep)

    def retitle_graph(self, uid, old_title, new_title, old_chart, new_chart):
        if old_title == new_title and old_chart == new_chart:
            return
        name = f"graphs/{uid}/sheet.json"
        sj = self.jget(name)
        sj["title"] = new_title
        self.jput(name, sj)
        bn = f"graphs/{uid}/data.bin"
        b = self.files[bn]
        b = self._splice_title(b, old_title, new_title, 0x0009, True)
        b = self._splice_title(b, old_chart, new_chart, 0x002C, False)
        self.files[bn] = b
        self.doc["sheetAttributesMap"].setdefault(uid, {})["title"] = new_title

    def retitle_data_sheet(self, uid, new_title):
        name = f"data/sheets/{uid}/sheet.json"
        sj = self.jget(name)
        if sj["title"] != new_title:
            sj["title"] = new_title
            self.jput(name, sj)
        self.doc["sheetAttributesMap"].setdefault(uid, {})["title"] = new_title

    def rename_in_json(self, names, old, new):
        """Replace `old` with `new` inside every string value of the named JSON files."""
        def walk(o):
            if isinstance(o, dict):
                return {k: walk(v) for k, v in o.items()}
            if isinstance(o, list):
                return [walk(v) for v in o]
            if isinstance(o, str) and old in o:
                return o.replace(old, new)
            return o
        for n in names:
            if n.endswith(".json"):
                self.jput(n, walk(self.jget(n)))

    # --------------------------------------------------------------------- rows
    def put_table(self, sheet_uid, lines):
        sj = self.jget(f"data/sheets/{sheet_uid}/sheet.json")
        t = sj["table"]
        csv_name = f"data/tables/{t['uid']}/data.csv"
        old = self.files[csv_name].decode("utf-8")
        text = "\n".join(lines) + ("\n" if old.endswith("\n") else "")
        self.files[csv_name] = text.encode("utf-8")
        n_rows, n_cols = len(lines), len(lines[0].split(","))
        cj_name = f"data/tables/{t['uid']}/content.json"
        cj = self.jget(cj_name)
        if n_cols != cj["numberOfColumns"]:
            raise ShapeError(
                f"'{sj['title']}' would have {n_cols} columns but the template's table has "
                f"{cj['numberOfColumns']}: the number of analytes differs from the template, "
                "which the template route cannot do (each graph binary holds one series per analyte). "
                "Build without a template instead: build --no-template (graphs are then yours to create)")
        if cj["numberOfRows"] != n_rows or cj["numberOfColumns"] != n_cols:
            cj["numberOfRows"], cj["numberOfColumns"] = n_rows, n_cols
            self.jput(cj_name, cj)
        sets = list(t["dataSets"]) + ([t["xDataSet"]] if t.get("xDataSet") else [])
        for s in sets:
            sn = f"data/sets/{s}.json"
            if sn not in self.files:
                continue
            sd = self.jget(sn)
            changed = False
            for rep in sd.get("replicates", []):
                if rep.get("lastRow") != n_rows - 1:
                    rep["lastRow"] = n_rows - 1
                    changed = True
            if changed:
                self.jput(sn, sd)
        sel = sj.get("selectionRange")
        if sel and sel["endCell"]["row"] > n_rows - 1:
            sel["endCell"]["row"] = n_rows - 1
            sel["startCell"]["row"] = min(sel["startCell"]["row"], n_rows - 1)
            self.jput(f"data/sheets/{sheet_uid}/sheet.json", sj)

    # ------------------------------------------------------------------ layouts
    def reconcile_layouts(self):
        for lid in list(self.doc["sheets"].get("layouts", [])):
            lj = self.jget(f"layouts/{lid}/sheet.json")
            gone = [g for g in lj.get("linkedGraphSheets", [])
                    if f"graphs/{g}/sheet.json" not in self.files]
            if gone:
                title = lj.get("title", lid)
                self._remove_files([n for n in self.files if n.startswith(f"layouts/{lid}/")])
                self._forget(lid)
                self.warnings.append(
                    f"layout '{title}' removed: {len(gone)} of its graphs no longer exist "
                    "(its placement records cannot be edited safely); recreate it in Prism "
                    "from the remaining graphs")

    def fix_current_sheets(self):
        cur = self.doc["uiSettings"].get("currentSheets", {})
        for key, lst in (("data", "data"), ("analysis", "analyses"), ("graph", "graphs"),
                         ("layout", "layouts"), ("info", "info")):
            if key in cur and cur[key] not in self.doc["sheets"].get(lst, []):
                rest = self.doc["sheets"].get(lst, [])
                if rest:
                    cur[key] = rest[0]
                else:
                    cur.pop(key)

    # --------------------------------------------------------------------- save
    def save(self, out_path):
        self.sync_doc()
        with zipfile.ZipFile(out_path, "w") as z:
            for name, data in self.files.items():
                z.writestr(self.infos[name], data, compress_type=self.infos[name].compress_type)
        return out_path

    # ----------------------------------------------------------------- validate
    def validate(self):
        problems = []
        ds, gs = self.data_sheets(), self.graph_sheets()
        lists = self.doc["sheets"]
        prefix = {"data": "data/sheets/", "graphs": "graphs/", "layouts": "layouts/",
                  "analyses": "analyses/", "info": "info/"}
        for key, lst in lists.items():
            for u in lst:
                if f"{prefix[key]}{u}/sheet.json" not in self.files:
                    problems.append(f"document.json lists {key} {u} but it has no files")
        for u in self.doc["sheetAttributesMap"]:
            if not any(u in lst for lst in lists.values()):
                if u not in ds:  # result sheets are not listed; others are orphans
                    problems.append(f"sheetAttributesMap has entry for unknown sheet {u}")
        all_sets = set()
        for u, sj in ds.items():
            t = sj["table"]
            all_sets |= set(t.get("dataSets", []))
            for k in ("xDataSet", "rowTitlesDataSet"):
                if t.get(k):
                    all_sets.add(t[k])
            for part in ("data.csv", "content.json", "data.dt"):
                if f"data/tables/{t['uid']}/{part}" not in self.files:
                    problems.append(f"table {t['uid']} ({sj['title']}) missing {part}")
        for u in lists["data"]:
            sj = ds.get(u)
            if not sj:
                continue
            t = sj["table"]
            try:
                rows = self.files[f"data/tables/{t['uid']}/data.csv"].decode().rstrip("\n").split("\n")
                cj = self.jget(f"data/tables/{t['uid']}/content.json")
                if len(rows) != cj["numberOfRows"] or len(rows[0].split(",")) != cj["numberOfColumns"]:
                    problems.append(f"{sj['title']}: csv is {len(rows)}x{len(rows[0].split(','))} "
                                    f"but content.json says {cj['numberOfRows']}x{cj['numberOfColumns']}")
            except KeyError:
                pass
        for g, gj in gs.items():
            for s in gj.get("inputDataSets", []):
                if s not in all_sets:
                    problems.append(f"graph '{gj['title']}' uses unknown dataset {s}")
        for a, aj in self.analyses().items():
            for s in aj.get("inputSheets", []):
                if s["uid"] not in ds:
                    problems.append(f"analysis '{aj['title']}' input sheet {s['uid']} missing")
            for n in [n for n in self.files if n.startswith(f"analyses/{a}/result_sheets/")
                      and n.endswith(".json")]:
                if self.jget(n)["dataSheet"] not in ds:
                    problems.append(f"analysis '{aj['title']}' result sheet table missing")
        for lid in lists.get("layouts", []):
            for g in self.jget(f"layouts/{lid}/sheet.json").get("linkedGraphSheets", []):
                if g not in gs:
                    problems.append(f"layout {lid} links missing graph {g}")
        # analyte counts must agree everywhere an analysis has one entry per analyte
        for a, aj in self.analyses().items():
            n = len(aj.get("inputDataSets", []))
            tag = f"analysis '{aj['title']}'"
            eq_sets = set()
            for sh in aj.get("inputSheets", []):
                if sh["uid"] in ds:
                    eq_sets |= set(ds[sh["uid"]]["table"]["dataSets"])
            for d in aj["inputDataSets"]:
                if d["uid"] not in eq_sets:
                    problems.append(f"{tag}: input dataset {d['title']!r} is not a column of its input sheet")
            par = self.jget(f"analyses/{a}/parameters.json")["content"]["model"]["initialValues"]
            for key, lst in par.items():
                if len(lst) != n:
                    problems.append(f"{tag}: {len(lst)} initial values for {key}, {n} analytes")
            res = self.jget(f"analyses/{a}/results.json")["content"]
            for key in ("models", "dataSummary"):
                if len(res[key]) != n:
                    problems.append(f"{tag}: {len(res[key])} entries in results {key}, {n} analytes")
            for rn in [x for x in self.files if x.startswith(f"analyses/{a}/result_sheets/") and x.endswith(".json")]:
                rs = self.jget(rn)
                t = ds[rs["dataSheet"]]["table"]
                if len(t["dataSets"]) != n:
                    problems.append(f"{tag}: result table '{rs['name']}' has {len(t['dataSets'])} datasets, {n} analytes")
        for u, sj in ds.items():                      # every table: csv shape == content.json
            t = sj["table"]
            try:
                raw = self.files[f"data/tables/{t['uid']}/data.csv"].decode()
                cj = self.jget(f"data/tables/{t['uid']}/content.json")
            except KeyError:
                continue
            if not raw.strip():
                continue
            rows = raw.rstrip("\n").split("\n")
            cols = {len(r.split(",")) for r in rows}
            if len(rows) != cj["numberOfRows"] or cols != {cj["numberOfColumns"]}:
                problems.append(f"table of '{sj['title']}': csv is {len(rows)} rows x {sorted(cols)} columns, "
                                f"content.json says {cj['numberOfRows']} x {cj['numberOfColumns']}")
            for sid in t.get("dataSets", []):
                f = f"data/sets/{sid}.json"
                if f not in self.files:
                    continue
        # orphan files
        owned_tables = {sj["table"]["uid"] for sj in ds.values()}
        for n in self.files:
            if n.startswith("data/tables/") and n.count("/") >= 3:
                if n.split("/")[2] not in owned_tables:
                    problems.append(f"orphan table file {n}")
        return problems


def apply_shape(project, ligands, conc_labels):
    """Reshape the project to `ligands` (names) x `conc_labels` (e.g. '5000 nM').

    Slot i of the template becomes ligand i; slot j of each ligand becomes concentration j.
    """
    blocks = project.ligand_blocks()
    if len(ligands) > len(blocks):
        raise ShapeError(
            f"{len(ligands)} ligands requested but the template holds {len(blocks)}. "
            "Adding ligands needs new graphs, which cannot be synthesized safely; use a "
            "template with at least that many ligands.")
    for b in blocks:
        if len(conc_labels) > len(b["kinetics"]):
            raise ShapeError(
                f"{len(conc_labels)} concentrations requested but the template holds "
                f"{len(b['kinetics'])} for '{b['ligand']}'. Adding concentrations needs new "
                "graphs, which cannot be synthesized safely; use a template with at least "
                "that many concentrations per ligand.")
    # surplus ligands out
    for b in blocks[len(ligands):]:
        for a in b["analysis"]:
            project.remove_analysis(a)
        if b["eq_graph"]:
            project.remove_graph(b["eq_graph"])
        for k, g in b["kinetics"]:
            project.remove_data_sheet(k)
            if g:
                project.remove_graph(g)
        project.remove_data_sheet(b["eq_sheet"])
    # rename / relabel the survivors
    for b, lig in zip(blocks, ligands):
        old = b["ligand"]
        # concentrations: drop surplus, retitle the rest
        for k, g in b["kinetics"][len(conc_labels):]:
            project.remove_data_sheet(k)
            if g:
                project.remove_graph(g)
        ds = project.data_sheets()
        for (k, g), label in zip(b["kinetics"], conc_labels):
            old_title = ds[k]["title"]
            old_label = old_title[len(f"Kinetics {old} - "):]
            new_title = f"Kinetics {lig} - {label}"
            project.retitle_data_sheet(k, new_title)
            if g:
                project.retitle_graph(g, old_title, new_title,
                                      f"{old} Kinetics - {old_label}",
                                      f"{lig} Kinetics - {label}")
        if lig != old:
            project.retitle_data_sheet(b["eq_sheet"], lig)
            if b["eq_graph"]:
                project.retitle_graph(b["eq_graph"], old, lig,
                                      f"{old} Equilibrium Binding", f"{lig} Equilibrium Binding")
            for a in b["analysis"]:
                names = [n for n in project.files if n.startswith(f"analyses/{a}/")]
                for n in [n for n in project.files if n.startswith(f"analyses/{a}/result_sheets/")
                          and n.endswith(".json")]:
                    rs = project.jget(n)["dataSheet"]
                    project.rename_in_json(project._sheet_files(rs), old, lig)
                    t = project.jget(f"data/sheets/{rs}/sheet.json")["title"]
                    project.doc["sheetAttributesMap"].setdefault(rs, {})["title"] = t
                project.rename_in_json(names, old, lig)
                project.doc["sheetAttributesMap"].setdefault(a, {})["title"] = \
                    project.jget(f"analyses/{a}/sheet.json")["title"]
    project.reconcile_layouts()
    project.fix_current_sheets()
    return project
