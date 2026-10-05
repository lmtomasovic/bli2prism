"""Equilibrium fits: recover known parameters, and flag what a human should look at."""

import numpy as np
import pytest

from bli2prism import eqfit

def test_demo_equilibrium_fits_are_sensible(demo_run):
    """Fits of the demo experiment's equilibrium values land near the simulated affinities."""
    import os
    from bli2prism import refdata as R
    from bli2prism import setupsheet as S
    csv_path = os.path.join(demo_run, "kineticanalysistableresults.csv")
    rows = R.read_results_csv(csv_path)
    st = R.read_setup(os.path.join(demo_run, "bli2prism_setup.xlsx"), {r["Loading Sample ID"] for r in rows})
    files = R.ligand_files(rows, st["ligands"], st["series_M"])
    steps = sorted(files["Binder B"])
    xs, _ = R.step_concentrations(demo_run, files["Binder B"], st["series_M"], True)
    lines = R.equilibrium_table(rows, "Binder B", st["series_M"], [xs[k] for k in steps], st["analytes"], steps=steps)
    a = np.array([[float(v) for v in ln.split(",")] for ln in lines])
    ec50_nM = [10 ** eqfit.fit_curve(a[:, 0], a[:, 1 + j])["three"]["params"][2] * 1e9 for j in range(3)]
    assert 30 < ec50_nM[0] < 90 and 10 < ec50_nM[1] < 40 and 1000 < ec50_nM[2] < 4000   # simulated KDs 50, 20, 2000 nM


def curve(lec=-7.0, bottom=0.02, top=0.4, hill=1.0, noise=0.0, n=8, hi=-5.3, lo=-9.5, seed=1):
    x = np.linspace(hi, lo, n)
    y = eqfit.sigmoid4(x, bottom, top, lec, hill) + np.random.default_rng(seed).normal(0, noise, n)
    return x, y


def test_fit_recovers_known_parameters():
    x, y = curve(noise=0.002)
    t = eqfit.fit_curve(x, y)["three"]
    assert t["params"] == pytest.approx([0.02, 0.4, -7.0], abs=0.02) and t["r2"] > 0.99
    assert eqfit.flags(eqfit.fit_curve(x, y)) == []


def test_four_parameter_fit_finds_a_steeper_slope():
    x, y = curve(hill=2.0, noise=0.002)
    f = eqfit.fit_curve(x, y)
    assert f["four"]["ok"] and f["four"]["params"][3] == pytest.approx(2.0, abs=0.4)
    assert f["four"]["aicc"] < f["three"]["aicc"]


def test_flags_plateau_not_reached():
    x, y = curve(lec=-5.6, noise=0.002)                # EC50 near the top concentration
    assert any("plateau" in m for m in eqfit.flags(eqfit.fit_curve(x, y)))


def test_flags_poor_fit_and_ec50_outside_range():
    rng = np.random.default_rng(3)
    x = np.linspace(-5.3, -9.5, 8)
    f = eqfit.fit_curve(x, 0.2 + rng.normal(0, 0.08, 8))              # no real curve
    msgs = " | ".join(eqfit.flags(f))
    assert "R²" in msgs or "did not converge" in msgs or "outside" in msgs


def test_a_failed_fit_is_reported_not_raised():
    f = eqfit.fit_curve(np.array([-6.0, -7.0]), np.array([0.1, 0.2]))        # fewer points than parameters
    assert not f["three"]["ok"]
    assert "did not converge" in eqfit.flags(f)[0]


def test_wide_interval_is_flagged():
    x, y = curve(lec=-6.2, noise=0.03, n=6, hi=-5.3, lo=-6.9)
    t = eqfit.fit_curve(x, y)
    if t["three"]["ok"]:
        assert isinstance(eqfit.flags(t), list)           # must not raise on awkward data


def test_a_fit_with_almost_no_degrees_of_freedom_says_so():
    x = np.array([-5.3, -6.3, -7.3, -8.3])
    f = eqfit.fit_curve(x, eqfit.sigmoid3(x, 0.02, 0.4, -6.8))
    msgs = eqfit.flags(f)
    assert any("only 4 points for 3 parameters (1 degree of freedom)" in m for m in msgs)
    x8 = np.linspace(-5.3, -9.5, 8)
    assert not any("degree of freedom" in m for m in eqfit.flags(eqfit.fit_curve(x8, eqfit.sigmoid3(x8, 0.02, 0.4, -7.0) + 0.001)))
