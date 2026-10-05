"""Equilibrium curve fits, done in Python so EC50 / KD is available without opening Prism.

fit3 reproduces Prism's 'Sigmoidal dose-response' (fixed slope):
    Y = Bottom + (Top - Bottom) / (1 + 10^(LogEC50 - X))      X = log10(concentration in M)
fit4 adds a Hill slope (Prism's 'variable slope' / 4PL), reported alongside for comparison;
the Prism template itself still applies the fixed-slope fit.

Confidence intervals here are the usual asymptotic (Wald) intervals from the fit covariance.
Prism reports profile-likelihood intervals, which are asymmetric and can differ noticeably
with only 8 points, so treat these as approximate.
"""

from __future__ import annotations

import math

import numpy as np
from scipy import stats
from scipy.optimize import curve_fit


def sigmoid3(x, b, t, lec):
    return b + (t - b) / (1 + 10 ** (lec - x))


def sigmoid4(x, b, t, lec, h):
    return b + (t - b) / (1 + 10 ** ((lec - x) * h))


def _fit(fn, names, x, y, p0):
    n, k = len(x), len(names)
    out = {"ok": False, "n": n, "k": k, "names": names}
    if n <= k:
        out["error"] = f"{n} point(s) cannot determine {k} parameters"
        return out
    try:
        with np.errstate(over="ignore", invalid="ignore"):
            popt, pcov = curve_fit(fn, x, y, p0=p0, maxfev=20000)
    except (RuntimeError, ValueError, TypeError) as e:
        out["error"] = str(e)
        return out
    resid = y - fn(x, *popt)
    sse = float(np.sum(resid ** 2))
    sst = float(np.sum((y - y.mean()) ** 2))
    df = n - k
    se = np.sqrt(np.diag(pcov)) if np.all(np.isfinite(pcov)) else np.full(k, np.nan)
    tcrit = stats.t.ppf(0.975, df) if df > 0 else float("nan")
    out.update(
        ok=True, params=[float(v) for v in popt], se=[float(v) for v in se],
        ci=[(float(p - tcrit * s), float(p + tcrit * s)) for p, s in zip(popt, se)],
        sse=sse, df=df, r2=1 - sse / sst if sst > 0 else float("nan"),
        syx=math.sqrt(sse / df) if df > 0 else float("nan"),
        aicc=(n * math.log(sse / n) + 2 * k + 2 * k * (k + 1) / (n - k - 1))
        if sse > 0 and n - k - 1 > 0 else float("nan"),
    )
    return out


def initial(x, y):
    mid = (y.min() + y.max()) / 2
    return [float(y.min()), float(y.max()), float(x[int(np.argmin(np.abs(y - mid)))])]


def fit_curve(x, y):
    """x = log10(M), y = response. -> {'three': fit, 'four': fit, 'x': x, 'y': y}."""
    x, y = np.asarray(x, float), np.asarray(y, float)
    p0 = initial(x, y)
    three = _fit(sigmoid3, ["Bottom", "Top", "LogEC50"], x, y, p0)
    four = _fit(sigmoid4, ["Bottom", "Top", "LogEC50", "Hill slope"], x, y,
                (three["params"] if three["ok"] else p0) + [1.0])
    return {"three": three, "four": four, "x": x, "y": y}


def flags(fit):
    """Plain-language problems with a 3-parameter fit; empty list = nothing to report."""
    t, x, y = fit["three"], fit["x"], fit["y"]
    if not t["ok"]:
        return [f"fit did not converge ({t.get('error', 'unknown')})"]
    b, top, lec = t["params"]
    out = []
    if t["df"] <= 1:
        out.append(f"only {t['n']} points for {t['k']} parameters ({t['df']} degree of freedom): the fit is barely "
                   "determined, so R² and the intervals mean little")
    if t["r2"] < 0.95:
        out.append(f"R² = {t['r2']:.3f} is below 0.95")
    if x.max() - lec < 1:
        out.append(f"top concentration is only {10 ** (x.max() - lec):.1f}× EC50: the top plateau "
                   "is not reached, so Top and EC50 are extrapolated")
    if lec < x.min() or lec > x.max():
        out.append("EC50 lies outside the tested concentration range")
    if top > 2 * y.max():
        out.append(f"fitted Top ({top:.3g}) is more than twice the highest response ({y.max():.3g})")
    lo, hi = t["ci"][2]
    if np.isfinite(lo) and np.isfinite(hi) and hi - lo > 1:
        out.append(f"LogEC50 95% interval spans {hi - lo:.2f} log units (approximate)")
    if not np.all(np.isfinite(t["se"])):
        out.append("standard errors could not be estimated")
    return out
