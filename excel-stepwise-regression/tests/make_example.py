"""Writes examples/Stepwise_Example_Output.xlsx: the add-in's report layout for
forward, backward and stepwise runs (t Stat criterion) on the Hoyda data.

Usage: python make_example.py path/to/Hoyda_Discrimination.xlsx output.xlsx
"""
import math
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from scipy import stats

import reference_stepwise as ref

F = Font(name="Arial", size=10)


def write_report(wb, ds, method, mname, sname):
    inm, log = ref.run(ds, method, ref.CRIT_T)
    ws = wb.create_sheet(sname)
    r = [3]

    def cell(row, col, v, font=F, align=None):
        c = ws.cell(row, col, v)
        c.font = font
        if align:
            c.alignment = Alignment(horizontal=align)
        return c

    def pair(label, v):
        cell(r[0], 1, label)
        cell(r[0], 2, v, align="left")
        r[0] += 1

    def header(hs, title):
        if title:
            cell(r[0], 1, title, Font(name="Arial", size=10, bold=True))
            r[0] += 1
        for i, h in enumerate(hs):
            cell(r[0], i + 1, h, Font(name="Arial", size=10, italic=True)).border = Border(bottom=Side(style="thin"))
        r[0] += 1

    def row(vals):
        for c, v in enumerate(vals):
            if v is not None:
                cell(r[0], c + 1, v)
        r[0] += 1

    cell(1, 1, "Stepwise Regression Output", Font(name="Arial", size=14, bold=True))
    pair("Dependent variable", "SALARY")
    pair("Method", mname)
    pair("Criterion", "t Stat")
    if method != 2:
        pair("|t| to enter", 2)
    if method != 1:
        pair("|t| to remove", 2)
    pair("Observations used", ds.n)
    pair("Rows excluded (blank or non-numeric)", 0)
    pair("Y range", "'[Hoyda Discrimination.xlsx]Reg - 5 preds'!$A$4:$A$48")
    pair("X range", "'[Hoyda Discrimination.xlsx]Reg - 5 preds'!$B$4:$F$48")
    r[0] += 1

    header(["Step", "Action", "Variable", "t Stat", "P-value", "Terms in model", "R Square",
            "Adjusted R Square", "RMSE", "AIC", "BIC"], "Selection steps")
    for i, (act, name, tv, pv, f) in enumerate(log):
        p, sse = len(f["vars"]), f["sse"]
        row([i, act, name or ("(all candidates)" if method == 2 else "(intercept only)"), tv, pv, p,
             1 - sse / ds.syy, 1 - (sse / f["dfe"]) / (ds.syy / (ds.n - 1)), math.sqrt(sse / f["dfe"]),
             ds.n * math.log(sse / ds.n) + 2 * (p + 1), ds.n * math.log(sse / ds.n) + (p + 1) * math.log(ds.n)])

    f = ref.current(ds, inm)
    p, sse, dfe = len(f["vars"]), f["sse"], f["dfe"]
    ssr = ds.syy - sse
    r[0] += 1
    header(["Regression Statistics", ""], "Final model")
    pair("Multiple R", math.sqrt(1 - sse / ds.syy))
    pair("R Square", 1 - sse / ds.syy)
    pair("Adjusted R Square", 1 - (sse / dfe) / (ds.syy / (ds.n - 1)))
    pair("RMSE", math.sqrt(sse / dfe))
    pair("Observations", ds.n)
    r[0] += 1

    header(["ANOVA", "df", "SS", "MS", "F", "Significance F"], "")
    fs = (ssr / p) / (sse / dfe)
    row(["Regression", p, ssr, ssr / p, fs, stats.f.sf(fs, p, dfe)])
    row(["Residual", dfe, sse, sse / dfe])
    row(["Total", ds.n - 1, ds.syy])
    r[0] += 1

    header(["Term", "Coefficients", "Standard Error", "t Stat", "P-value", "Lower 95%", "Upper 95%"], "")
    tc = stats.t.ppf(0.975, dfe)
    for i, nm in enumerate(["Intercept"] + [ds.names[j] for j in f["vars"]]):
        b, s = f["coef"][i], f["se"][i]
        row([nm, b, s, b / s, ref.pval(b, s, dfe), b - tc * s, b + tc * s])
    r[0] += 1

    header(["Variable", "t Stat if entered next", "P-value if entered next", "Note"],
           "Variables not in the final model")
    for j in range(ds.k):
        if not inm[j]:
            inm[j] = True
            t = ref.current(ds, inm)
            inm[j] = False
            pos = t["vars"].index(j) + 1
            row([ds.names[j], ref.tstat(t["coef"][pos], t["se"][pos]),
                 ref.pval(t["coef"][pos], t["se"][pos], t["dfe"])])

    for col, w in zip("ABCDEFGHIJK", [34, 22, 22, 12, 12, 15, 12, 18, 10, 10, 10]):
        ws.column_dimensions[col].width = w


if __name__ == "__main__":
    ds = ref.load(sys.argv[1], "Reg - 5 preds", "A", "B:F", 4, 48)
    wb = Workbook()
    wb.remove(wb.active)
    write_report(wb, ds, 1, "Forward selection", "Stepwise")
    write_report(wb, ds, 2, "Backward elimination", "Stepwise (2)")
    write_report(wb, ds, 3, "Stepwise", "Stepwise (3)")
    wb.save(sys.argv[2])
