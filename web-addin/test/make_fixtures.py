"""Generates fixtures.json: synthetic datasets with expected results from the
Python reference port (../../excel-stepwise-regression/tests) and scipy.

Usage: REFERENCE_DIR=path/to/excel-stepwise-regression/tests python make_fixtures.py [Hoyda.xlsx]
The optional Hoyda workbook adds a local-only hoyda.json (not committed).
"""
import json
import os
import sys

import numpy as np
from scipy import stats

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.environ.get("REFERENCE_DIR", os.path.join(HERE, "..", "..", "excel-stepwise-regression", "tests")))
import reference_stepwise as ref  # noqa: E402


def expected(ds):
    out = {}
    for crit in (1, 2):
        for method in (1, 2, 3):
            try:
                inm, log = ref.run(ds, method, crit)
            except Exception:
                continue
            if method == 2 and ref.fit(ds, list(range(ds.k))) is None:
                out[f"{method}-{crit}"] = {"error": True}
                continue
            f = ref.current(ds, inm)
            out[f"{method}-{crit}"] = {
                "chosen": [ds.names[j] for j in f["vars"]],
                "coef": f["coef"], "se": f["se"], "sse": f["sse"],
                "steps": [[a, n, t, p] for a, n, t, p, _ in log],
            }
    return out


def dataset(seed, n=60, k=6):
    rng = np.random.default_rng(seed)
    X = rng.normal(size=(n, k)) * rng.uniform(0.5, 30, size=k) + rng.uniform(-500, 500, size=k)
    X[:, 2] = X[:, 0] * 3 - X[:, 1]
    y = 10 + 0.6 * X[:, 0] - 0.2 * X[:, 3] + rng.normal(scale=15, size=n)
    names = [f"x{j}" for j in range(k)]
    return y.tolist(), X.T.tolist(), names


fixtures = {"datasets": [], "dist": []}
for seed in range(8):
    y, cols, names = dataset(seed)
    ds = ref.DataSet(y, [list(r) for r in zip(*cols)], names)
    fixtures["datasets"].append({"y": y, "cols": cols, "names": names, "expected": expected(ds)})

for t, df in [(0.5, 3), (2.0, 10), (2.27, 40), (5.04, 42), (28.3, 40), (0.01, 1), (3.5, 200)]:
    fixtures["dist"].append({"kind": "t2", "x": t, "df": df, "p": float(2 * stats.t.sf(t, df))})
for f, d1, d2 in [(18.58, 3, 40), (1.2, 2, 10), (0.3, 5, 38), (10.83, 5, 38)]:
    fixtures["dist"].append({"kind": "f", "x": f, "d1": d1, "d2": d2, "p": float(stats.f.sf(f, d1, d2))})
for df in [1, 5, 40, 1000]:
    fixtures["dist"].append({"kind": "tinv", "df": df, "t": float(stats.t.ppf(0.975, df))})

with open(os.path.join(HERE, "fixtures.json"), "w") as fh:
    json.dump(fixtures, fh)

if len(sys.argv) > 1:
    ds = ref.load(sys.argv[1], "Reg - 5 preds", "A", "B:F", 4, 48)
    cols = [[r[j] for r in ds.x] for j in range(ds.k)]
    with open(os.path.join(HERE, "hoyda.json"), "w") as fh:
        json.dump({"y": ds.y, "cols": cols, "names": ds.names, "expected": expected(ds)}, fh)
