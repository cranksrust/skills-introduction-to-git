# Stepwise Regression add-in for Excel

A VBA add-in (`.xlam`) for desktop Excel on Windows and Mac. It runs forward selection, backward elimination, or stepwise regression and writes a report in the same layout as the Analysis ToolPak's Regression output, plus the selection history.

Excel for the web and Excel on iPad do not run VBA, so the add-in does not work there.

## Install

### Option A: build script (Windows)

1. In Excel, enable **File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model**.
2. Run `powershell -ExecutionPolicy Bypass -File .\Build-Addin.ps1` in this folder. It writes `StepwiseRegression.xlam`.

### Option B: by hand (Windows or Mac)

1. Open a blank workbook and press **Alt+F11** (Mac: **Tools > Macro > Visual Basic Editor**).
2. **File > Import File...** and pick `StepwiseRegression.bas`.
3. Close the editor, then **File > Save As**, type **Excel Add-in (\*.xlam)**. Excel suggests its AddIns folder.

### Load it

**File > Options > Add-ins > Manage: Excel Add-ins > Go...**, tick **StepwiseRegression** (use **Browse** if it is not listed). On Mac: **Tools > Excel Add-ins**. Restart Excel once so the right-click menu item appears.

## Use

### Interactive report

Right-click any cell and choose **Stepwise Regression...** (or **Alt+F8**, type `RunStepwiseRegression`, **Run**). The prompts ask for:

| Prompt | Notes |
|---|---|
| Y range | One column. Include the header cell if you have labels. |
| X range | Candidate predictors. Ctrl+click to pick non-adjacent columns. Each must cover the same rows as Y. |
| Labels | Whether the first row is headers. |
| Method | 1 Forward, 2 Backward, 3 Stepwise. |
| P to enter / P to remove | Defaults 0.05 / 0.10. Stepwise requires enter <= remove. |

The report goes on a new sheet named `Stepwise` in the data's workbook:

- **Selection steps:** each entry or removal with its p-value, R², adjusted R², standard error, AIC and BIC.
- **Final model:** regression statistics, ANOVA, and coefficients with standard errors, t, p and 95% intervals.
- **Variables not in the final model:** the p-value each would have if entered next, or a note when it is constant or collinear.

### Worksheet function

```
=STEPREG(known_y, known_x, [method], [p_enter], [p_remove], [has_labels])
```

Returns the final coefficient table (Term, Coefficient, Std Error, t Stat, P-value). It spills in Excel 365. In older versions select a block 5 columns wide and enough rows, type the formula and press **Ctrl+Shift+Enter**. `known_x` must be one contiguous block. Returns `#VALUE!` on bad input.

## Method

- Entry and removal use the partial t-test on the candidate's coefficient, which is equivalent to the partial F-test.
- Forward: start with the intercept only, add the candidate with the smallest p-value while it is below p-to-enter.
- Backward: start with every candidate, drop the term with the largest p-value while it is above p-to-remove.
- Stepwise: one forward step, then backward removals until none qualify, repeated until nothing changes. A 1000-iteration cap guards against cycling and is flagged in the report.
- Rows where Y or any X is blank, text or an error are excluded. The report states how many.
- Least squares is solved on mean-centred cross products. Candidates that are constant or collinear with terms already in the model (1 - R² below 1e-10) are skipped.

## Worked example: Hoyda salary data

On the `Reg - 5 preds` sheet of `Hoyda_Discrimination.xlsx`, Y = `A4:A48`, X = `B4:F48`, labels on, default thresholds. All three methods choose the same model:

| Step | Forward / Stepwise | Backward |
|---|---|---|
| 1 | SENIORITY enters (p = 9.2e-6) | SEX_SENIOR removed (p = 0.794) |
| 2 | ROW4 enters (p = 0.00079) | SEX removed (p = 0.511) |
| 3 | SENIORITY2 enters (p = 0.0288) | |

| Term | Coefficient | Std Error | P-value |
|---|---|---|---|
| Intercept | 576.851 | 20.349 | 4.2e-28 |
| SENIORITY | 18.367 | 4.523 | 0.00022 |
| SENIORITY2 | -0.4554 | 0.2008 | 0.0288 |
| ROW4 | -157.480 | 41.967 | 0.00056 |

R² = 0.5822, n = 44. SEX does not enter: once seniority and the row-4 teacher are accounted for, sex adds nothing significant.

## Verification

`tests/reference_stepwise.py` is a line-for-line Python port of the VBA algorithm. `tests/test_reference.py` checks it against statsmodels on synthetic data (including an exactly collinear predictor) and against the Hoyda workbook's own full-model output.

```
pip install numpy scipy statsmodels openpyxl pytest
HOYDA_XLSX=path/to/Hoyda_Discrimination.xlsx python -m pytest tests
```
