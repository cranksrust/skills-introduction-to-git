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
3. Double-click **ThisWorkbook** in the Project Explorer and paste:
   ```vb
   Private Sub Workbook_Open()
       Auto_Open
   End Sub

   Private Sub Workbook_BeforeClose(Cancel As Boolean)
       Auto_Close
   End Sub
   ```
   This adds the right-click menu item each time Excel loads the add-in.
4. Close the editor, then **File > Save As**, type **Excel Add-in (\*.xlam)**. Excel suggests its AddIns folder.

### Load it

**File > Options > Add-ins > Manage: Excel Add-ins > Go...**, tick **StepwiseRegression** (use **Browse** if it is not listed). On Mac: **Tools > Excel Add-ins**. Restart Excel once so the right-click menu item appears.

### Add a Ribbon button (recommended)

Newer Excel builds can rebuild the right-click menu after add-ins load, so the menu item may not appear. A Ribbon button always works:

1. **File > Options > Customize Ribbon**.
2. On the right, select **Data**, click **New Group**, then **Rename** it to `Stepwise`.
3. On the left, set **Choose commands from** to **Macros**, select **RunStepwiseRegression**, click **Add >>**, then **OK**.

## Use

### Interactive report

Right-click any cell and choose **Stepwise Regression...** (or **Alt+F8**, type `RunStepwiseRegression`, **Run**). The prompts ask for:

| Prompt | Notes |
|---|---|
| Y range | One column. Include the header cell if you have labels. |
| X range | Candidate predictors. Ctrl+click to pick non-adjacent columns. Each must cover the same rows as Y. |
| Labels | Whether the first row is headers. |
| Method | 1 Forward, 2 Backward, 3 Stepwise. |
| Criterion | 1 t Stat (default), 2 P-value. |
| Enter / remove thresholds | t Stat: enter when \|t\| >= 2.0, remove when \|t\| < 2.0. P-value: enter when p < 0.05, remove when p > 0.10. Stepwise requires \|t\| to enter >= \|t\| to remove (or p to enter <= p to remove). |

The report goes on a new sheet named `Stepwise` in the data's workbook:

- **Selection steps:** each entry or removal with its t Stat and p-value, plus R², adjusted R², RMSE, AIC and BIC after the step.
- **Final model:** R, R², adjusted R², RMSE, ANOVA, and coefficients with standard errors, t, p and 95% intervals.
- **Variables not in the final model:** the t Stat and p-value each would have if entered next, or a note when it is constant or collinear.

### Worksheet function

```
=STEPREG(known_y, known_x, [method], [criterion], [enter_threshold], [remove_threshold], [has_labels])
```

`method` 1/2/3 (default 3), `criterion` 1 t Stat (default) or 2 P-value. Thresholds default to 2.0/2.0 for t Stat and 0.05/0.10 for P-value. Example: `=STEPREG(A4:A48, B4:F48, 3, 1, 2, 2, TRUE)`.

Returns the final coefficient table (Term, Coefficient, Std Error, t Stat, P-value). It spills in Excel 365. In older versions select a block 5 columns wide and enough rows, type the formula and press **Ctrl+Shift+Enter**. `known_x` must be one contiguous block. Returns `#VALUE!` on bad input.

## Method

- Entry and removal use the partial t-test on the candidate's coefficient, which is equivalent to the partial F-test. The t Stat criterion compares \|t\| with a fixed cutoff, the P-value criterion compares the two-tailed p-value.
- Forward: start with the intercept only, add the candidate with the largest \|t\| while it qualifies.
- Backward: start with every candidate, drop the term with the smallest \|t\| while it fails the remove threshold.
- Stepwise: one forward step, then backward removals until none qualify, repeated until nothing changes. A 1000-iteration cap guards against cycling and is flagged in the report.
- Rows where Y or any X is blank, text or an error are excluded. The report states how many.
- Least squares is solved on mean-centred cross products. Candidates that are constant or collinear with terms already in the model (1 - R² below 1e-10) are skipped.

## Worked example: Hoyda salary data

On the `Reg - 5 preds` sheet of `Hoyda_Discrimination.xlsx`, Y = `A4:A48`, X = `B4:F48`, labels on, default thresholds. All three methods, under either criterion, choose the same model. Full reports are in `examples/Stepwise_Example_Output.xlsx`.

| Step | Forward / Stepwise | Backward |
|---|---|---|
| 1 | SENIORITY enters (t = 5.04) | SEX_SENIOR removed (t = 0.26) |
| 2 | ROW4 enters (t = -3.63) | SEX removed (t = -0.66) |
| 3 | SENIORITY2 enters (t = -2.27) | |

| Term | Coefficient | Std Error | t Stat | P-value |
|---|---|---|---|---|
| Intercept | 576.851 | 20.349 | 28.35 | 4.2e-28 |
| SENIORITY | 18.367 | 4.523 | 4.06 | 0.00022 |
| SENIORITY2 | -0.4554 | 0.2008 | -2.27 | 0.0288 |
| ROW4 | -157.480 | 41.967 | -3.75 | 0.00056 |

R² = 0.5822, RMSE = 39.88, n = 44. SEX does not enter: once seniority and the row-4 teacher are accounted for, sex adds nothing significant.

## Verification

`tests/reference_stepwise.py` is a line-for-line Python port of the VBA algorithm. `tests/test_reference.py` checks it against statsmodels on synthetic data (including an exactly collinear predictor) and against the Hoyda workbook's own full-model output.

```
pip install numpy scipy statsmodels openpyxl pytest
HOYDA_XLSX=path/to/Hoyda_Discrimination.xlsx python -m pytest tests
```
