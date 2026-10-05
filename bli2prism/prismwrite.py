"""Clone a template .prism and replace its data tables / sheet titles.

A .prism is a zip of JSON plus per-table data.csv. Only the files named in
`replacements` change; every other member (graphs, analyses, layouts, ...) is copied
through byte-for-byte, in the original member order.
"""

from __future__ import annotations

import json
import os
import zipfile


def table_index(zin):
    """title -> table uid, for every data sheet in the project."""
    out = {}
    for name in zin.namelist():
        if name.startswith("data/sheets/") and name.endswith("/sheet.json"):
            d = json.loads(zin.read(name))
            out[d["title"]] = d["table"]["uid"]
    return out


def clone_with_tables(template_path, out_path, table_csv_by_title):
    """Write out_path = template with data.csv replaced for each {sheet title: [lines]}."""
    with zipfile.ZipFile(template_path) as zin:
        idx = table_index(zin)
        missing = [t for t in table_csv_by_title if t not in idx]
        if missing:
            raise KeyError(
                f"template has no data sheet titled {missing}. Cloning only replaces "
                "existing sheets; adding sheets for a different shape (more concentrations "
                "than the template) is not implemented yet.")
        repl = {f"data/tables/{idx[t]}/data.csv": "\n".join(lines) + "\n"
                for t, lines in table_csv_by_title.items()}
        # keep the original trailing-newline convention per file
        for name in repl:
            old = zin.read(name).decode("utf-8")
            if not old.endswith("\n"):
                repl[name] = repl[name].rstrip("\n")
        with zipfile.ZipFile(out_path, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename in repl:
                    data = repl[info.filename].encode("utf-8")
                zout.writestr(info, data, compress_type=info.compress_type)
    return out_path


def describe_project(path):
    """What a .prism holds, read from its document.json. 'is_blank' is true for a project as
    Prism creates it (at most its one placeholder data sheet, no analyses): the kind to upload
    as an empty container. Raises ValueError if the file is not a Prism project."""
    try:
        with zipfile.ZipFile(path) as z:
            doc = json.loads(z.read("document.json"))
    except (zipfile.BadZipFile, KeyError, json.JSONDecodeError) as e:
        raise ValueError(f"{os.path.basename(path)} is not a Prism project ({type(e).__name__})") from e
    sheets = doc.get("sheets", {})
    n = {k: len(sheets.get(k, [])) for k in ("data", "graphs", "layouts", "analyses")}
    return {**n, "is_blank": n["data"] <= 1 and n["analyses"] == 0}


def diff_projects(a_path, b_path):
    """Members that differ between two .prism files (name, reason)."""
    out = []
    with zipfile.ZipFile(a_path) as a, zipfile.ZipFile(b_path) as b:
        an, bn = a.namelist(), b.namelist()
        for n in sorted(set(an) ^ set(bn)):
            out.append((n, "only in " + ("A" if n in an else "B")))
        for n in an:
            if n in bn and a.read(n) != b.read(n):
                out.append((n, "content differs"))
        if an != bn:
            out.append(("(member order)", "differs"))
    return out
