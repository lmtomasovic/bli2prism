"""Kinetic fit-row selection: the rules, and the demo run's known outcome."""

import os

import pytest

from bli2prism import kinetics as K


def row(**kw):
    base = {"Full  X^2": "0.05", "Full  R^2": "0.98", "KD (M)": "1.0E-08", "kon(1/Ms)": "1.0E+06",
            "kdis(1/s)": "1.0E-02", "KD2": "<1.0E-12", "kon2": "1.0E+05", "kdis2": "1.0E-02"}
    base.update(kw)
    return base


def test_a_good_row_is_used():
    c = K.classify(row())
    assert c["status"] == "used" and c["kd_M"] == 1e-8 and c["reason"] == ""


@pytest.mark.parametrize("kw,fragment", [
    ({"Full  X^2": "3"}, "X²"),                       # not strictly below 3
    ({"Full  X^2": "5.2"}, "X²"),
    ({"Full  R^2": "0.9"}, "R²"),                     # not strictly above 0.9
    ({"Full  R^2": "0.2"}, "R²"),
    ({"KD (M)": "<1.0E-12"}, "KD censored"),
    ({"kon(1/Ms)": "4.20E+99"}, "kon did not converge"),
    ({"kdis(1/s)": "<1.0E-07"}, "kdis censored"),
    ({"kon(1/Ms)": "<1.0E+03"}, "kon did not converge"),
])
def test_exclusion_rules(kw, fragment):
    c = K.classify(row(**kw))
    assert c["status"] == "excluded" and fragment in c["reason"]


def test_boundary_values_just_inside_are_used():
    assert K.classify(row(**{"Full  X^2": "2.999", "Full  R^2": "0.9001"}))["status"] == "used"


def test_kd2_is_never_used():
    """Decision: KD values only. A censored KD with a determinate KD2 stays excluded."""
    c = K.classify(row(**{"KD (M)": "<1.0E-12", "KD2": "7.01E-08", "kon2": "1.24E+06", "kdis2": "8.66E-02"}))
    assert c["status"] == "excluded" and c["kd_M"] is None
    assert "second site" in c["reason"] and "KD2" in c["reason"]       # noted for information only
    assert not hasattr(K, "kd2_fallback")
    with pytest.raises(TypeError):
        K.classify(row(), True)                                         # no fallback switch exists


def test_failed_filters_win_over_kd2_note():
    c = K.classify(row(**{"Full  R^2": "0.5", "KD (M)": "<1.0E-12", "KD2": "7.01E-08"}))
    assert "R²" in c["reason"] and "second site" not in c["reason"]


def _group(kds, **kw):
    rows = []
    for i, kd in enumerate(kds):
        rows.append({"Loading Sample ID": "L", "Sample ID": "A", "Conc. (nM)": str(100 * (i + 1)),
                     **row(**{"KD (M)": f"{kd:.3E}", **kw})})
    return K.build_table(rows, ["L"], ["A"])


def test_outlier_flag_needs_three_used_rows_and_a_five_fold_gap():
    t = _group([1e-8, 1.2e-8, 9e-9, 2e-7])
    assert [c["flagged"] for c in t] == [False, False, False, True]
    assert "still included" in t[3]["reason"] and t[3]["status"] == "used"
    assert not any(c["flagged"] for c in _group([1e-8, 1.2e-8, 4e-8]))            # 4x: not an outlier
    assert not any(c["flagged"] for c in _group([1e-8, 5e-7]))                    # only two rows: no judgement


def test_a_minority_of_outliers_cannot_make_good_rows_look_bad():
    """Two outliers among five: only they are flagged (the median resists them)."""
    t = _group([1e-8, 1e-8, 1e-8, 1e-5, 1e-5])
    assert [c["flagged"] for c in t] == [False, False, False, True, True]


def test_the_flag_reason_names_the_fold_and_the_median():
    t = _group([1e-8, 1.2e-8, 9e-9, 2e-7])
    assert "18" in t[3]["reason"] and "median" in t[3]["reason"] and "4 used rows" in t[3]["reason"]


def test_demo_run_outcome():
    import os
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "examples", "demo_binder_screen", "kineticanalysistableresults.csv")
    t = K.build_table(K.read_fit_rows(path), ["Binder A", "Binder B"], ["Analyte X", "Analyte Y", "Analyte Z"])
    assert len(t) == 48

    def kds(lig, a):
        return [c["kd_M"] * 1e9 for c in t if c["ligand"] == lig and c["analyte"] == a and c["status"] == "used"]
    assert kds("Binder A", "Analyte Z") == []                           # never binds: every row censored
    assert all(c["status"] == "excluded" for c in t if c["ligand"] == "Binder A" and c["analyte"] == "Analyte Z")
    b = kds("Binder B", "Analyte X")
    assert len(b) >= 4 and 40 < sum(b) / len(b) < 60                    # simulated KD 50 nM
    assert not any(c["flagged"] for c in t)


def test_csv_duplicate_headers_use_the_first_column(tmp_path):
    p = tmp_path / "r.csv"
    p.write_text("Sample ID,Full  X^2,Full  X^2\nS,0.05,9.9\n")
    assert K.read_fit_rows(str(p))[0]["Full  X^2"] == "0.05"           # first of the two identical columns
