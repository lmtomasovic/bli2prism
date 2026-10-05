"""The formula evaluator: checked against Excel's own cached results, then unit by unit."""

import math

import openpyxl
import pytest

from bli2prism import xlcalc as X



def make(cells, sheet2=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for k, v in cells.items():
        ws[k] = v
    if sheet2:
        w2 = wb.create_sheet("T")
        for k, v in sheet2.items():
            w2[k] = v
    return wb


def ev(formula, cells=None, sheet2=None):
    wb = make({"Z1": formula, **(cells or {})}, sheet2)
    return X.evaluate_workbook(wb)["S"][(1, 26)]


@pytest.mark.parametrize("formula,expected", [
    ("=1+2*3", 7), ("=(1+2)*3", 9), ("=2^3^2", 64), ("=-2^2", 4), ("=10/4", 2.5),
    ('="a"&"b"', "ab"), ('=1&"x"', "1x"), ("=1<2", True), ('="A"="a"', True), ("=3<>3", False),
    ("=LOG(1000)", 3), ("=LOG10(100)", 2), ("=ABS(-3)", 3), ("=ROUND(2.567,1)", 2.6),
    ("=SUM(1,2,3)", 6), ("=MAX(1,5,3)", 5), ("=MIN(4,2)", 2), ("=OR(1>2,2>1)", True),
    ("=ISNUMBER(3)", True), ('=ISNUMBER("x")', False),
])
def test_scalar_formulas(formula, expected):
    assert ev(formula) == pytest.approx(expected) if not isinstance(expected, (str, bool)) else ev(formula) == expected


def test_ranges_and_blank_handling():
    cells = {"A1": 1, "A2": 2, "A3": "text", "A4": None, "A5": 4}
    assert ev("=SUM(A1:A5)", cells) == 7 and ev("=AVERAGE(A1:A5)", cells) == pytest.approx(7 / 3)
    assert ev("=MAX(A1:A5)", cells) == 4 and ev("=A4+1", cells) == 1


def test_conditional_aggregates():
    cells = {"A1": "x", "A2": "y", "A3": "x", "A4": "x", "B1": 10, "B2": 20, "B3": 30, "B4": "n/a",
             "C1": "Yes", "C2": "Yes", "C3": "No", "C4": "Yes"}
    assert ev('=COUNTIFS(A1:A4,"x")', cells) == 3
    assert ev('=COUNTIFS(A1:A4,"x",C1:C4,"Yes")', cells) == 2
    assert ev('=AVERAGEIFS(B1:B4,A1:A4,"x")', cells) == 20          # text cells ignored
    assert ev('=AVERAGEIFS(B1:B4,A1:A4,"x",C1:C4,"<>No")', cells) == 10
    assert ev('=SUMIF(A1:A4,"x",B1:B4)', cells) == 40
    assert ev('=COUNTIFS(B1:B4,">15")', cells) == 2


def test_blank_criterion_matches_empty_cells():
    cells = {"A1": "k", "A2": None, "A3": "k", "B1": 1, "B2": 2, "B3": 3}
    assert ev('=COUNTIFS(A1:A3,"")', cells) == 1


def test_if_and_iferror_are_lazy():
    assert ev('=IF(1>2,1/0,"ok")') == "ok"
    assert ev("=IFERROR(1/0,\"n/a\")") == "n/a"
    assert ev('=IFERROR(AVERAGEIFS(A1:A2,B1:B2,"z"),"none")', {"A1": 1, "A2": 2, "B1": "a", "B2": "b"}) == "none"
    assert ev('=IF(ISNUMBER(A1),IF(A1>2,"big","small"),"")', {"A1": 5}) == "big"


def test_cross_sheet_references():
    assert ev("=T!A1*2+S!A1", {"A1": 1}, {"A1": 5}) == 11
    assert ev("='T'!A1&\"-\"", None, {"A1": "q"}) == "q-"
    assert ev('=SUMIF(T!A1:A3,"x",T!B1:B3)', None, {"A1": "x", "A2": "y", "A3": "x", "B1": 1, "B2": 2, "B3": 4}) == 5


def test_nested_formula_cells_do_not_clobber_each_other():
    cells = {"A1": 2, "A2": "=A1*3", "A3": "=A2+A1", "Z2": "=A3*A2+A1"}
    wb = make({"Z1": "=A3*A2+A1", **cells})
    assert X.evaluate_workbook(wb)["S"][(1, 26)] == 8 * 6 + 2


def test_division_by_zero_becomes_an_excel_error_value():
    assert ev("=1/0") == "#DIV/0!"


def test_unsupported_functions_raise_instead_of_being_guessed():
    wb = make({"Z1": '=VLOOKUP(1,A1:B2,2,0)'})
    with pytest.raises(X.XlError, match="unsupported function VLOOKUP"):
        X.Evaluator(wb).value("Z1", "S")


def test_function_names_with_digits_are_not_cell_references():
    assert ev("=LOG10(100)") == 2 and ev("=LOG10(1000)+A1", {"A1": 1}) == 4


def test_circular_references_are_detected():
    wb = make({"A1": "=A2", "A2": "=A1"})
    with pytest.raises(X.XlError, match="circular"):
        X.Evaluator(wb).value("A1", "S")


def test_cached_values_are_written_into_the_file(tmp_path):
    wb = make({"A1": 2, "A2": "=A1*5", "A3": '=IF(A1>5,"big","small")', "A4": "=1/0"})
    path = str(tmp_path / "c.xlsx")
    wb.save(path)
    X.fill_cached_values(path, wb)
    v = openpyxl.load_workbook(path, data_only=True)["S"]
    assert (v["A2"].value, v["A3"].value, v["A4"].value) == (10, "small", "#DIV/0!")
    assert openpyxl.load_workbook(path)["S"]["A2"].value == "=A1*5"          # formulas kept


def test_lookup_and_text_functions():
    import openpyxl
    from bli2prism import xlcalc as X
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for i, v in enumerate(["x", "y", "z"], 1):
        ws[f"A{i}"] = v
    for ref, f in (("C1", '=INDEX(A1:A3,MATCH("y",A1:A3,0))'), ("C2", '=IFERROR(MATCH("q",A1:A3,0),"none")'),
                   ("C3", '=MID("ABCDEFGH",3,1)&MID("ABCDEFGH",5,2)'), ("C4", '=IF(AND(1=1,NOT(2=3)),"yes","no")'),
                   ("C5", '=IFERROR(INDEX(A1:A3,4),"out")'), ("C6", '=MATCH("Y",A1:A3,0)')):
        ws[ref] = f
    ev = X.Evaluator(wb)
    assert [ev.value(f"C{i}", "S") for i in range(1, 7)] == ["y", "none", "CEF", "yes", "out", 2]
