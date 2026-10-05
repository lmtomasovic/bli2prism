"""Minimal Excel formula evaluator, used to verify the workbook and to store cached values.

openpyxl writes formulas without results, so a file opened by anything other than Excel
(Quick Look, pandas, a previewer) shows empty cells. `fill_cached_values` evaluates every
formula with this evaluator and writes the result into the file as the cached value; Excel
still recalculates on open.

Supports (also across sheets, as 'Sheet'!A1): numbers, strings, A1 / $A$1 refs and ranges on the same sheet, + - * / ^ & and
comparisons, and SUM AVERAGE MAX MIN COUNTIFS AVERAGEIFS SUMIF IF IFERROR ISNUMBER OR ROUND SUMIFS.
LOG LOG10 ABS AND NOT MID INDEX MATCH (exact). Anything else raises, so an unsupported formula can never be silently mis-evaluated.
"""

from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from xml.etree import ElementTree as ET

from openpyxl.utils import column_index_from_string, get_column_letter

TOKEN = re.compile(r"""\s*(?:
    (?P<num>\d+\.?\d*(?:[eE][+-]?\d+)?)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<func>[A-Z][A-Z0-9.]*)\(
  | (?P<range>(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_]*)!)?\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)
  | (?P<ref>(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_]*)!)?\$?[A-Z]{1,3}\$?\d+)
  | (?P<op><>|<=|>=|[-+*/^&=<>(),])
)""", re.X)


class XlError(Exception):
    pass


def _tokens(src):
    pos, out = 0, []
    while pos < len(src):
        m = TOKEN.match(src, pos)
        if not m or m.end() == pos:
            raise XlError(f"cannot tokenize {src[pos:pos+20]!r} in {src!r}")
        pos = m.end()
        kind = m.lastgroup
        out.append((kind, m.group(kind)))
    return out


def _cell(ref):
    m = re.match(r"\$?([A-Z]+)\$?(\d+)", ref)
    return int(m.group(2)), column_index_from_string(m.group(1))


def _split(ref):
    """'Sheet'!A1 -> ('Sheet', 'A1'); A1 -> (None, 'A1')."""
    if "!" in ref:
        sheet, addr = ref.rsplit("!", 1)
        return sheet.strip("'"), addr
    return None, ref


class Evaluator:
    def __init__(self, wb):
        self.wb, self.cache, self.stack = wb, {}, set()
        self.sheet = None

    # --- cell access
    def value(self, ref, sheet=None):
        sh, addr = _split(ref)
        sheet = sh or sheet or self.sheet
        r, c = _cell(addr)
        key = (sheet, r, c)
        if key in self.cache:
            return self.cache[key]
        if key in self.stack:
            raise XlError(f"circular reference at {sheet}!{addr}")
        v = self.wb[sheet].cell(r, c).value
        if isinstance(v, str) and v.startswith("="):
            self.stack.add(key)
            prev, self.sheet = self.sheet, sheet
            try:
                v = self.eval(v)
            finally:
                self.sheet = prev
                self.stack.discard(key)
        self.cache[key] = v
        return v

    def rng(self, text):
        sh, addr = _split(text)
        sheet = sh or self.sheet
        a, b = addr.split(":")
        (r0, c0), (r1, c1) = _cell(a), _cell(b)
        return [[self.value(f"{get_column_letter(c)}{r}", sheet) for c in range(c0, c1 + 1)]
                for r in range(r0, r1 + 1)]

    def eval(self, formula):
        # Evaluating a reference to another formula cell re-enters here mid-parse, so the
        # parser state must be saved and restored.
        saved = (getattr(self, "toks", None), getattr(self, "i", 0))
        self.toks, self.i = _tokens(formula.lstrip("=")), 0
        try:
            v = self._cmp()
            if self.i != len(self.toks):
                raise XlError(f"unparsed input in {formula!r}")
            return v
        finally:
            self.toks, self.i = saved

    # --- parser
    def _peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def _take(self):
        t = self.toks[self.i]
        self.i += 1
        return t

    def _cmp(self):
        left = self._cat()
        k, v = self._peek()
        if k == "op" and v in ("=", "<>", "<", ">", "<=", ">="):
            self._take()
            right = self._cat()
            return _compare(left, right, v)
        return left

    def _cat(self):
        v = self._add()
        while self._peek() == ("op", "&"):
            self._take()
            v = f"{_text(v)}{_text(self._add())}"
        return v

    def _add(self):
        v = self._mul()
        while self._peek()[0] == "op" and self._peek()[1] in "+-":
            op = self._take()[1]
            r = self._mul()
            v = _num(v) + _num(r) if op == "+" else _num(v) - _num(r)
        return v

    def _mul(self):
        v = self._pow()
        while self._peek()[0] == "op" and self._peek()[1] in "*/":
            op = self._take()[1]
            r = self._pow()
            if op == "*":
                v = _num(v) * _num(r)
            else:
                if _num(r) == 0:
                    raise ZeroDivisionError
                v = _num(v) / _num(r)
        return v

    def _pow(self):
        v = self._unary()
        while self._peek() == ("op", "^"):
            self._take()
            v = _num(v) ** _num(self._unary())
        return v

    def _unary(self):
        if self._peek() == ("op", "-"):
            self._take()
            return -_num(self._unary())
        if self._peek() == ("op", "+"):
            self._take()
            return self._unary()
        return self._primary()

    def _primary(self):
        kind, val = self._take()
        if kind == "num":
            return float(val)
        if kind == "str":
            return val[1:-1].replace('""', '"')
        if kind == "ref":
            return self.value(val)
        if kind == "range":
            return self.rng(val)
        if kind == "op" and val == "(":
            v = self._cmp()
            assert self._take() == ("op", ")")
            return v
        if kind == "func":
            return self._call(val)
        raise XlError(f"unexpected token {kind} {val}")

    def _args(self, lazy_first=False):
        args = []
        if self._peek() == ("op", ")"):
            self._take()
            return args
        while True:
            args.append(self._cmp())
            k, v = self._take()
            if v == ")":
                return args

    def _call(self, name):
        if name in ("IF", "IFERROR"):           # lazy: only the taken branch may be valid
            return self._lazy(name)
        args = self._args()
        flat = lambda a: [x for row in a for x in row] if isinstance(a, list) else [a]
        nums = lambda a: [x for x in flat(a) if isinstance(x, (int, float)) and not isinstance(x, bool)]
        if name == "SUM":
            return sum(sum(nums(a)) for a in args)
        if name == "AVERAGE":
            xs = [x for a in args for x in nums(a)]
            if not xs:
                raise ZeroDivisionError
            return sum(xs) / len(xs)
        if name == "MAX":
            return max(x for a in args for x in nums(a))
        if name == "MIN":
            return min(x for a in args for x in nums(a))
        if name == "ISNUMBER":
            return isinstance(args[0], (int, float)) and not isinstance(args[0], bool)
        if name == "OR":
            return any(bool(x) for a in args for x in flat(a))
        if name == "AND":
            return all(bool(x) for a in args for x in flat(a))
        if name == "NOT":
            return not bool(args[0])
        if name == "MID":
            s, start, n = _text(args[0]), int(_num(args[1])), int(_num(args[2]))
            return s[start - 1:start - 1 + n]
        if name == "MATCH":                      # exact match only (match type 0)
            if len(args) < 3 or _num(args[2]) != 0:
                raise XlError("MATCH is supported with match type 0 only")
            for i, x in enumerate(flat(args[1]), 1):
                if x is not None and x != "" and _compare(x, args[0], "="):
                    return i
            raise XlError("#N/A")
        if name == "INDEX":                      # one row or one column of values, INDEX(range, n)
            items = flat(args[0])
            n = int(_num(args[1]))
            if not 1 <= n <= len(items):
                raise XlError("#REF!")
            return items[n - 1]
        if name in ("LOG", "LOG10"):
            import math
            return math.log10(_num(args[0]))
        if name == "ABS":
            return abs(_num(args[0]))
        if name == "ROUND":
            return round(_num(args[0]), int(_num(args[1])))
        if name == "SUMIF":
            rg, crit = flat(args[0]), args[1]
            tgt = flat(args[2]) if len(args) > 2 else rg
            return sum(t for c, t in zip(rg, tgt)
                       if _match(c, crit) and isinstance(t, (int, float)) and not isinstance(t, bool))
        if name == "SUMIFS":
            target, pairs = flat(args[0]), args[1:]
            ranges = [flat(pairs[i]) for i in range(0, len(pairs), 2)]
            crits = [pairs[i] for i in range(1, len(pairs), 2)]
            return sum(target[j] for j in range(len(ranges[0]))
                       if all(_match(rg[j], cr) for rg, cr in zip(ranges, crits))
                       and isinstance(target[j], (int, float)) and not isinstance(target[j], bool))
        if name in ("COUNTIFS", "AVERAGEIFS"):
            if name == "AVERAGEIFS":
                target, pairs = flat(args[0]), args[1:]
            else:
                pairs, target = args, None
            ranges = [flat(pairs[i]) for i in range(0, len(pairs), 2)]
            crits = [pairs[i] for i in range(1, len(pairs), 2)]
            hits = [j for j in range(len(ranges[0]))
                    if all(_match(rg[j], cr) for rg, cr in zip(ranges, crits))]
            if name == "COUNTIFS":
                return len(hits)
            xs = [target[j] for j in hits
                  if isinstance(target[j], (int, float)) and not isinstance(target[j], bool)]
            if not xs:
                raise ZeroDivisionError
            return sum(xs) / len(xs)
        raise XlError(f"unsupported function {name}")

    def _lazy(self, name):
        # Evaluate arguments with exception capture: parse all, but only the chosen branch's
        # errors matter. Re-parse by index to skip over unneeded branches.
        start = self.i
        results = []
        depth, j = 0, self.i
        spans, begin = [], self.i
        while True:
            k, v = self.toks[j]
            if k == "func" or v == "(":
                depth += 1
            elif v == ")":
                if depth == 0:
                    spans.append((begin, j))
                    break
                depth -= 1
            elif v == "," and depth == 0:
                spans.append((begin, j))
                begin = j + 1
            j += 1
        end = j + 1

        def run(span):
            self.i = span[0]
            sub_end = span[1]
            saved = self.toks
            self.toks = saved[:sub_end]
            try:
                return self._cmp()
            finally:
                self.toks = saved

        try:
            if name == "IF":
                cond = run(spans[0])
                out = run(spans[1] if cond else spans[2]) if len(spans) > 2 or cond else False
            else:  # IFERROR
                try:
                    out = run(spans[0])
                except (ZeroDivisionError, ValueError, ArithmeticError, XlError):
                    out = run(spans[1])
        finally:
            self.i = end
        return out


def _num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return v
    raise ValueError(f"#VALUE! {v!r}")


def _text(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return "" if v is None else str(v)


def _compare(a, b, op):
    if a is None:
        a = "" if isinstance(b, str) else 0
    if b is None:
        b = "" if isinstance(a, str) else 0
    if isinstance(a, str) and isinstance(b, str):
        a, b = a.lower(), b.lower()
    elif isinstance(a, str) != isinstance(b, str):
        return {"=": False, "<>": True}.get(op, isinstance(a, str) and op in (">", ">="))
    return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op]


def _match(cell, crit):
    """COUNTIFS/AVERAGEIFS criterion semantics (no wildcards)."""
    if isinstance(crit, str):
        m = re.match(r"(<>|<=|>=|<|>|=)?(.*)$", crit, re.S)
        op, rest = m.group(1) or "=", m.group(2)
        if rest == "":
            blank = cell is None or cell == ""
            return blank if op == "=" else not blank
        try:
            rest = float(rest)
        except ValueError:
            pass
        if isinstance(rest, str):
            c = "" if cell is None else str(cell)
            return (c.lower() == rest.lower()) == (op == "=") if op in ("=", "<>") else False
        crit, cell_cmp = rest, cell
        if not isinstance(cell_cmp, (int, float)) or isinstance(cell_cmp, bool):
            return op == "<>"
        return _compare(cell_cmp, crit, op)
    if isinstance(cell, (int, float)) and not isinstance(cell, bool):
        return cell == crit
    return False


# ---------------------------------------------------------------- cached values
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def evaluate_workbook(wb):
    """{sheet title: {(row, col): value}} for every formula cell. Errors become '#DIV/0!' etc."""
    out, ev = {}, Evaluator(wb)
    for ws in wb.worksheets:
        res = {}
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    try:
                        res[(c.row, c.column)] = ev.value(c.coordinate, ws.title)
                    except ZeroDivisionError:
                        res[(c.row, c.column)] = "#DIV/0!"
                    except ValueError:
                        res[(c.row, c.column)] = "#VALUE!"
        out[ws.title] = res
    return out


def fill_cached_values(xlsx_path, wb):
    """Rewrite xlsx_path so every formula cell carries its computed value."""
    values = evaluate_workbook(wb)
    sheet_files = {}
    with zipfile.ZipFile(xlsx_path) as z:
        wbxml = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid = {r.get("Id"): r.get("Target") for r in rels}
        for s in wbxml.find(f"{{{NS}}}sheets"):
            target = rid[s.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")]
            sheet_files["xl/" + target.lstrip("/").replace("xl/", "")] = s.get("name")
    tmp = tempfile.mktemp(suffix=".xlsx")
    ET.register_namespace("", NS)
    with zipfile.ZipFile(xlsx_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in sheet_files:
                vals = values[sheet_files[item.filename]]
                root = ET.fromstring(data)
                for c in root.iter(f"{{{NS}}}c"):
                    f = c.find(f"{{{NS}}}f")
                    if f is None:
                        continue
                    r, col = _cell(c.get("r"))
                    v = vals[(r, col)]
                    ve = c.find(f"{{{NS}}}v")
                    if ve is None:
                        ve = ET.SubElement(c, f"{{{NS}}}v")
                    if isinstance(v, bool):
                        c.set("t", "b")
                        ve.text = "1" if v else "0"
                    elif isinstance(v, (int, float)):
                        c.attrib.pop("t", None)
                        ve.text = repr(float(v))
                    elif isinstance(v, str) and v.startswith("#"):
                        c.set("t", "e")
                        ve.text = v
                    else:
                        c.set("t", "str")
                        ve.text = "" if v is None else str(v)
                data = ET.tostring(root, xml_declaration=True, encoding="UTF-8")
            zout.writestr(item, data)
    shutil.move(tmp, xlsx_path)
    return values
