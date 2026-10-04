"""Checks the reference port against statsmodels and the Hoyda workbook's own regression output.

Run: HOYDA_XLSX=path/to/Hoyda_Discrimination.xlsx python -m pytest tests
"""
import os

import numpy as np
import pytest
import statsmodels.api as sm

import reference_stepwise as ref

HOYDA = os.environ.get("HOYDA_XLSX")


def ols(ds, cols):
    X = sm.add_constant(np.array([[r[j] for j in cols] for r in ds.x]), has_constant="add")
    return sm.OLS(np.array(ds.y), X).fit()


@pytest.mark.skipif(not HOYDA, reason="set HOYDA_XLSX")
def test_hoyda_full_model_matches_workbook():
    ds = ref.load(HOYDA, "Reg - 5 preds", "A", "B:F", 4, 48)
    f = ref.fit(ds, list(range(5)))
    sheet_coef = [581.7649005931399, -21.171385924088554, 17.697192575902562,
                  -0.43399450105575177, -158.96189371905257, 1.3494421214570025]
    sheet_se = [22.521987347147608, 37.37388120399218, 4.880356284387182,
                0.2120258540533571, 70.29223513550541, 5.14407564298496]
    assert np.allclose(f["coef"], sheet_coef, rtol=1e-9)
    assert np.allclose(f["se"], sheet_se, rtol=1e-9)
    assert f["sse"] == pytest.approx(62788.607850916684, rel=1e-9)


@pytest.mark.skipif(not HOYDA, reason="set HOYDA_XLSX")
@pytest.mark.parametrize("crit", [ref.CRIT_T, ref.CRIT_P])
@pytest.mark.parametrize("method", [1, 2, 3])
def test_hoyda_selection(method, crit):
    ds = ref.load(HOYDA, "Reg - 5 preds", "A", "B:F", 4, 48)
    inm, log = ref.run(ds, method, crit)
    chosen = [ds.names[j] for j in range(ds.k) if inm[j]]
    assert chosen == ["SENIORITY", "SENIORITY2", "ROW4"]
    f = ref.current(ds, inm)
    res = ols(ds, f["vars"])
    assert np.allclose(f["coef"], res.params, rtol=1e-9)
    assert np.allclose(f["se"], res.bse, rtol=1e-9)


def synthetic(seed, n=80, k=8):
    rng = np.random.default_rng(seed)
    X = rng.normal(size=(n, k)) * rng.uniform(0.5, 50, size=k) + rng.uniform(-1000, 1000, size=k)
    X[:, 3] = X[:, 1] * 2 - X[:, 2]                 # exact collinearity
    y = 5 + 0.8 * X[:, 0] - 0.05 * X[:, 4] + rng.normal(scale=10, size=n)
    return ref.DataSet(list(y), X.tolist(), [f"x{j}" for j in range(k)])


def brute_forward(ds, p_enter=None, t_enter=None):
    """Independent forward selection with statsmodels."""
    chosen = []
    while True:
        best = None
        for j in range(ds.k):
            if j in chosen:
                continue
            cols = sorted(chosen + [j])
            if np.linalg.matrix_rank(np.array([[r[c] for c in cols] for r in ds.x]) -
                                     np.array([[ds.xbar[c] for c in cols]])) < len(cols):
                continue
            res = ols(ds, cols)
            pv = res.pvalues[cols.index(j) + 1]
            t = abs(res.tvalues[cols.index(j) + 1])
            if best is None or pv < best[1]:
                best = (j, pv, t)
        if best is None:
            return sorted(chosen)
        if p_enter is not None and best[1] >= p_enter:
            return sorted(chosen)
        if t_enter is not None and best[2] < t_enter:
            return sorted(chosen)
        chosen.append(best[0])


@pytest.mark.parametrize("crit", [ref.CRIT_T, ref.CRIT_P])
@pytest.mark.parametrize("seed", range(10))
def test_forward_matches_statsmodels(seed, crit):
    ds = synthetic(seed)
    inm, _ = ref.run(ds, 1, crit)
    expected = brute_forward(ds, t_enter=2.0) if crit == ref.CRIT_T else brute_forward(ds, p_enter=0.05)
    assert [j for j in range(ds.k) if inm[j]] == expected
    f = ref.current(ds, inm)
    res = ols(ds, f["vars"])
    assert np.allclose(f["coef"], res.params, rtol=1e-8)
    assert np.allclose(f["se"], res.bse, rtol=1e-8)


def test_collinear_full_model_rejected():
    ds = synthetic(0)
    assert ref.fit(ds, list(range(ds.k))) is None
