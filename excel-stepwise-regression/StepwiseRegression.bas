Attribute VB_Name = "StepwiseRegression"
'==============================================================================
' Stepwise Regression add-in for Microsoft Excel (Windows and Mac desktop)
'
' Entry points
'   RunStepwiseRegression   Interactive macro. Also on the right-click cell menu
'                           when the add-in is loaded.
'   =STEPREG(known_y, known_x, [method], [criterion], [enter_threshold], [remove_threshold], [has_labels])
'                           Worksheet function returning the final coefficient
'                           table (spills in Excel 365, array-enter elsewhere).
'
' Methods
'   1 = Forward selection     start empty, add the best candidate while it qualifies
'   2 = Backward elimination  start full, drop the weakest term while it fails
'   3 = Stepwise              forward step, then backward removals, repeat
'
' Criteria
'   1 = t Stat (default)      enter when |t| >= enter (2.0), remove when |t| < remove (2.0)
'   2 = P-value               enter when p < enter (0.05), remove when p > remove (0.10)
'
' Both criteria use the partial t-test on the candidate's coefficient (equivalent
' to the partial F-test), and the report shows t and p for every step.
' Rows where Y or any X is blank or non-numeric are excluded (listwise).
'==============================================================================
Option Explicit

Private Const APP_TITLE As String = "Stepwise Regression | HOUMBA '28"
Private Const CREDIT As String = "Built for the HOUMBA Class of '28"

' Report colours (Long values of RGB(r, g, b); Const cannot call RGB)
Private Const BURNT_ORANGE As Long = 22463      ' RGB(191, 87, 0)  #BF5700
Private Const ORANGE_TINT As Long = 14083576    ' RGB(248, 229, 214)
Private Const GREY_TEXT As Long = 5855577       ' RGB(89, 89, 89)

' Report number formats
Private Const FMT_P As String = "[<0.0001]""<0.0001"";0.0000"
Private Const FMT_T As String = "0.00"
Private Const FMT_R2 As String = "0.0000"
Private Const FMT_NUM As String = "#,##0.00"
Private Const FMT_COEF As String = "#,##0.0000"
Private Const MENU_TAG As String = "StepwiseRegressionAddin"
Private Const SING_TOL As Double = 1E-10
Private Const MAX_STEPS As Long = 1000

Private Const METHOD_FORWARD As Long = 1
Private Const METHOD_BACKWARD As Long = 2
Private Const METHOD_STEPWISE As Long = 3

Private Const CRIT_T As Long = 1
Private Const CRIT_P As Long = 2

Private Type DataSet
    n As Long
    k As Long
    y() As Double          ' 1..n
    x() As Double          ' 1..n, 1..k
    yBar As Double
    syy As Double          ' centered sum of squares of y
    xBar() As Double       ' 1..k
    cxx() As Double        ' 1..k, 1..k centered cross products
    cxy() As Double        ' 1..k
    xNames() As String     ' 1..k
    yName As String
End Type

Private Type FitResult
    ok As Boolean
    p As Long              ' predictors in model, excluding intercept
    vars() As Long         ' 1..p column indexes into the data set
    coef() As Double       ' 0..p, 0 = intercept
    se() As Double         ' 0..p
    sse As Double
    dfe As Long
End Type

Private Type StepLog
    nSteps As Long
    action() As String
    varName() As String
    tValue() As Double     ' t Stat of the variable entered or removed
    pValue() As Double     ' -1 when not applicable
    nVars() As Long
    r2() As Double
    adjR2() As Double
    s() As Double
    aic() As Double
    bic() As Double
    hitLimit As Boolean
End Type

'------------------------------------------------------------------------------
' Add-in load / unload: right-click menu entry
'------------------------------------------------------------------------------
Public Sub Auto_Open()
    AddMenu
End Sub

Public Sub Auto_Close()
    RemoveMenu
End Sub

Private Sub AddMenu()
    Dim ctl As Object
    On Error Resume Next
    RemoveMenu
    Set ctl = Application.CommandBars("Cell").Controls.Add(Type:=1, Temporary:=True)
    ctl.Caption = "Stepwise Regression..."
    ctl.Tag = MENU_TAG
    ctl.BeginGroup = True
    ctl.OnAction = "'" & ThisWorkbook.Name & "'!RunStepwiseRegression"
End Sub

Private Sub RemoveMenu()
    Dim cb As Object, i As Long
    On Error Resume Next
    Set cb = Application.CommandBars("Cell")
    For i = cb.Controls.Count To 1 Step -1
        If cb.Controls(i).Tag = MENU_TAG Then cb.Controls(i).Delete
    Next i
End Sub

'------------------------------------------------------------------------------
' Interactive macro
'------------------------------------------------------------------------------
Public Sub RunStepwiseRegression()
    Dim yRng As Range, xRng As Range, a As Range
    Dim v As Variant, hasLabels As Boolean, method As Long
    Dim crit As Long, enterVal As Double, removeVal As Double
    Dim ds As DataSet, inModel() As Boolean, sl As StepLog
    Dim dropped As Long, msg As String

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open the workbook that holds your data first.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    On Error Resume Next
    Set yRng = Application.InputBox( _
        "Select the dependent variable (Y): one column, including the label cell if you have one.", _
        APP_TITLE, Type:=8)
    On Error GoTo 0
    If yRng Is Nothing Then Exit Sub
    If yRng.Areas.Count > 1 Or yRng.Columns.Count <> 1 Then
        MsgBox "Y must be a single contiguous column.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    On Error Resume Next
    Set xRng = Application.InputBox( _
        "Select the candidate predictor columns (X). Hold Ctrl to add non-adjacent columns." & vbLf & _
        "Every column must cover the same rows as Y.", APP_TITLE, Type:=8)
    On Error GoTo 0
    If xRng Is Nothing Then Exit Sub
    For Each a In xRng.Areas
        If a.Rows.Count <> yRng.Rows.Count Or a.Row <> yRng.Row Then
            MsgBox "Every X column must start and end on the same rows as Y (" & _
                yRng.Address(False, False) & ").", vbExclamation, APP_TITLE
            Exit Sub
        End If
    Next a

    v = MsgBox("Does the first row of your selections contain labels?", _
        vbYesNoCancel + vbQuestion, APP_TITLE)
    If v = vbCancel Then Exit Sub
    hasLabels = (v = vbYes)

    v = Application.InputBox("Method:" & vbLf & _
        "  1 = Forward selection" & vbLf & _
        "  2 = Backward elimination" & vbLf & _
        "  3 = Stepwise (forward with backward removal)", APP_TITLE, 3, Type:=1)
    If VarType(v) = vbBoolean Then Exit Sub
    method = CLng(v)
    If method < 1 Or method > 3 Then
        MsgBox "Enter 1, 2 or 3.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    v = Application.InputBox("Criterion for entering and removing variables:" & vbLf & _
        "  1 = t Stat (|t| threshold)" & vbLf & _
        "  2 = P-value", APP_TITLE, CRIT_T, Type:=1)
    If VarType(v) = vbBoolean Then Exit Sub
    crit = CLng(v)
    If crit <> CRIT_T And crit <> CRIT_P Then
        MsgBox "Enter 1 or 2.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    DefaultThresholds crit, enterVal, removeVal

    If method <> METHOD_BACKWARD Then
        If crit = CRIT_T Then
            v = Application.InputBox("|t| to enter (a candidate enters when its |t Stat| is at least this):", _
                APP_TITLE, enterVal, Type:=1)
        Else
            v = Application.InputBox("P-value to enter (a candidate enters when its p-value is below this):", _
                APP_TITLE, enterVal, Type:=1)
        End If
        If VarType(v) = vbBoolean Then Exit Sub
        enterVal = CDbl(v)
    End If
    If method <> METHOD_FORWARD Then
        If crit = CRIT_T Then
            v = Application.InputBox("|t| to remove (a term is dropped when its |t Stat| is below this):", _
                APP_TITLE, IIf(method = METHOD_STEPWISE, Application.WorksheetFunction.Min(removeVal, enterVal), removeVal), Type:=1)
        Else
            v = Application.InputBox("P-value to remove (a term is dropped when its p-value is above this):", _
                APP_TITLE, IIf(method = METHOD_STEPWISE, Application.WorksheetFunction.Max(removeVal, enterVal), removeVal), Type:=1)
        End If
        If VarType(v) = vbBoolean Then Exit Sub
        removeVal = CDbl(v)
    End If
    msg = CheckThresholds(method, crit, enterVal, removeVal)
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, APP_TITLE
        Exit Sub
    End If

    If Not BuildDataSet(yRng, xRng, hasLabels, ds, dropped, msg) Then
        MsgBox msg, vbExclamation, APP_TITLE
        Exit Sub
    End If
    If Not RunSelection(ds, method, crit, enterVal, removeVal, inModel, sl, msg) Then
        MsgBox msg, vbExclamation, APP_TITLE
        Exit Sub
    End If

    On Error GoTo WriteFailed
    Application.ScreenUpdating = False
    WriteReport ds, method, crit, enterVal, removeVal, inModel, sl, dropped, yRng, xRng
    Application.ScreenUpdating = True
    Exit Sub

WriteFailed:
    Application.ScreenUpdating = True
    MsgBox "Could not write the report: " & Err.Description, vbExclamation, APP_TITLE
End Sub

'------------------------------------------------------------------------------
' Worksheet function
'------------------------------------------------------------------------------
Public Function STEPREG(known_y As Range, known_x As Range, Optional method As Long = 3, _
        Optional criterion As Long = 1, Optional enter_threshold As Variant, Optional remove_threshold As Variant, _
        Optional has_labels As Boolean = False) As Variant
    Dim ds As DataSet, inModel() As Boolean, sl As StepLog, f As FitResult
    Dim dropped As Long, msg As String, outArr() As Variant, i As Long, a As Range
    Dim enterVal As Double, removeVal As Double

    On Error GoTo Fail
    If method < 1 Or method > 3 Then GoTo Fail
    If criterion <> CRIT_T And criterion <> CRIT_P Then GoTo Fail
    DefaultThresholds criterion, enterVal, removeVal
    If Not IsMissing(enter_threshold) Then enterVal = CDbl(enter_threshold)
    If Not IsMissing(remove_threshold) Then removeVal = CDbl(remove_threshold)
    If Len(CheckThresholds(method, criterion, enterVal, removeVal)) > 0 Then GoTo Fail
    If known_y.Areas.Count > 1 Or known_y.Columns.Count <> 1 Then GoTo Fail
    For Each a In known_x.Areas
        If a.Rows.Count <> known_y.Rows.Count Then GoTo Fail
    Next a
    If Not BuildDataSet(known_y, known_x, has_labels, ds, dropped, msg) Then GoTo Fail
    If Not RunSelection(ds, method, criterion, enterVal, removeVal, inModel, sl, msg) Then GoTo Fail

    f = FitCurrent(ds, inModel)
    ReDim outArr(1 To f.p + 2, 1 To 5)
    outArr(1, 1) = "Term"
    outArr(1, 2) = "Coefficient"
    outArr(1, 3) = "Std Error"
    outArr(1, 4) = "t Stat"
    outArr(1, 5) = "P-value"
    For i = 0 To f.p
        If i = 0 Then outArr(2, 1) = "Intercept" Else outArr(i + 2, 1) = ds.xNames(f.vars(i))
        outArr(i + 2, 2) = f.coef(i)
        outArr(i + 2, 3) = f.se(i)
        outArr(i + 2, 4) = TStat(f.coef(i), f.se(i))
        outArr(i + 2, 5) = PValueT(f.coef(i), f.se(i), f.dfe)
    Next i
    STEPREG = outArr
    Exit Function

Fail:
    STEPREG = CVErr(xlErrValue)
End Function

'------------------------------------------------------------------------------
' Criteria
'------------------------------------------------------------------------------
Private Sub DefaultThresholds(crit As Long, enterVal As Double, removeVal As Double)
    If crit = CRIT_T Then
        enterVal = 2
        removeVal = 2
    Else
        enterVal = 0.05
        removeVal = 0.1
    End If
End Sub

' Returns an error message, or "" when the thresholds are usable
Private Function CheckThresholds(method As Long, crit As Long, enterVal As Double, removeVal As Double) As String
    If crit = CRIT_T Then
        If enterVal <= 0 Or removeVal <= 0 Then
            CheckThresholds = "|t| thresholds must be greater than 0."
        ElseIf method = METHOD_STEPWISE And enterVal < removeVal Then
            CheckThresholds = "|t| to enter must be at least |t| to remove, otherwise a variable can cycle in and out."
        End If
    Else
        If enterVal <= 0 Or enterVal >= 1 Or removeVal <= 0 Or removeVal >= 1 Then
            CheckThresholds = "P-value thresholds must be between 0 and 1."
        ElseIf method = METHOD_STEPWISE And enterVal > removeVal Then
            CheckThresholds = "P-value to enter must not exceed p-value to remove, otherwise a variable can cycle in and out."
        End If
    End If
End Function

Private Function EnterOK(crit As Long, t As Double, pv As Double, enterVal As Double) As Boolean
    If crit = CRIT_T Then EnterOK = (t >= enterVal) Else EnterOK = (pv < enterVal)
End Function

Private Function RemoveOK(crit As Long, t As Double, pv As Double, removeVal As Double) As Boolean
    If crit = CRIT_T Then RemoveOK = (t < removeVal) Else RemoveOK = (pv > removeVal)
End Function

'------------------------------------------------------------------------------
' Data preparation
'------------------------------------------------------------------------------
Private Function BuildDataSet(yRng As Range, xRng As Range, hasLabels As Boolean, _
        ds As DataSet, dropped As Long, msg As String) As Boolean
    Dim yv As Variant, cols() As Variant, a As Range, ov As Range
    Dim nRows As Long, firstRow As Long, k As Long, j As Long, c As Long
    Dim r As Long, i As Long, n As Long, useRow() As Boolean
    Dim sy As Double, sx As Double, dy As Double, dj As Double

    On Error Resume Next
    Set ov = Application.Intersect(yRng, xRng)
    On Error GoTo 0
    If Not ov Is Nothing Then
        msg = "The Y column must not be part of the X selection."
        Exit Function
    End If

    nRows = yRng.Rows.Count
    firstRow = IIf(hasLabels, 2, 1)
    yv = ColumnValues(yRng)

    k = 0
    For Each a In xRng.Areas
        k = k + a.Columns.Count
    Next a
    ReDim cols(1 To k)
    ReDim ds.xNames(1 To k)
    j = 0
    For Each a In xRng.Areas
        For c = 1 To a.Columns.Count
            j = j + 1
            cols(j) = ColumnValues(a.Columns(c))
            If hasLabels Then
                ds.xNames(j) = LabelText(cols(j)(1, 1), "Col " & ColLetter(a.Columns(c)))
            Else
                ds.xNames(j) = "Col " & ColLetter(a.Columns(c))
            End If
        Next c
    Next a
    If hasLabels Then
        ds.yName = LabelText(yv(1, 1), "Y")
    Else
        ds.yName = "Col " & ColLetter(yRng)
    End If

    ReDim useRow(1 To nRows)
    n = 0
    For r = firstRow To nRows
        useRow(r) = (VarType(yv(r, 1)) = vbDouble)
        If useRow(r) Then
            For j = 1 To k
                If VarType(cols(j)(r, 1)) <> vbDouble Then
                    useRow(r) = False
                    Exit For
                End If
            Next j
        End If
        If useRow(r) Then n = n + 1
    Next r
    dropped = (nRows - firstRow + 1) - n

    If n < 3 Then
        msg = "Fewer than 3 complete numeric rows. Check the ranges and the labels setting."
        Exit Function
    End If

    ds.n = n
    ds.k = k
    ReDim ds.y(1 To n)
    ReDim ds.x(1 To n, 1 To k)
    i = 0
    For r = firstRow To nRows
        If useRow(r) Then
            i = i + 1
            ds.y(i) = yv(r, 1)
            For j = 1 To k
                ds.x(i, j) = cols(j)(r, 1)
            Next j
        End If
    Next r

    ' Means, then centered cross products (two-pass for accuracy)
    ReDim ds.xBar(1 To k)
    ReDim ds.cxx(1 To k, 1 To k)
    ReDim ds.cxy(1 To k)
    sy = 0
    For i = 1 To n
        sy = sy + ds.y(i)
    Next i
    ds.yBar = sy / n
    For j = 1 To k
        sx = 0
        For i = 1 To n
            sx = sx + ds.x(i, j)
        Next i
        ds.xBar(j) = sx / n
    Next j
    ds.syy = 0
    For i = 1 To n
        dy = ds.y(i) - ds.yBar
        ds.syy = ds.syy + dy * dy
        For j = 1 To k
            dj = ds.x(i, j) - ds.xBar(j)
            ds.cxy(j) = ds.cxy(j) + dj * dy
            For c = j To k
                ds.cxx(j, c) = ds.cxx(j, c) + dj * (ds.x(i, c) - ds.xBar(c))
            Next c
        Next j
    Next i
    For j = 1 To k
        For c = j + 1 To k
            ds.cxx(c, j) = ds.cxx(j, c)
        Next c
    Next j

    If ds.syy <= 0 Then
        msg = "Y is constant across the usable rows, so there is nothing to explain."
        Exit Function
    End If
    BuildDataSet = True
End Function

Private Function ColumnValues(rng As Range) As Variant
    Dim v As Variant, arr() As Variant
    v = rng.Value2
    If IsArray(v) Then
        ColumnValues = v
    Else
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = v
        ColumnValues = arr
    End If
End Function

Private Function LabelText(v As Variant, fallback As String) As String
    If IsError(v) Or IsEmpty(v) Then
        LabelText = fallback
    Else
        LabelText = Trim$(CStr(v))
        If Len(LabelText) = 0 Then LabelText = fallback
    End If
End Function

Private Function ColLetter(rng As Range) As String
    Dim s As String
    s = rng.Cells(1, 1).Address(False, False)
    Do While Len(s) > 0 And IsNumeric(Right$(s, 1))
        s = Left$(s, Len(s) - 1)
    Loop
    ColLetter = s
End Function

'------------------------------------------------------------------------------
' Selection procedure
'------------------------------------------------------------------------------
Private Function RunSelection(ds As DataSet, method As Long, crit As Long, enterVal As Double, _
        removeVal As Double, inModel() As Boolean, sl As StepLog, msg As String) As Boolean
    Dim j As Long, changed As Boolean, guard As Long, f As FitResult

    ReDim inModel(1 To ds.k)
    sl.nSteps = 0
    sl.hitLimit = False

    If method = METHOD_BACKWARD Then
        If ds.n < ds.k + 2 Then
            msg = "Backward elimination needs at least " & (ds.k + 2) & " complete rows for " & _
                ds.k & " predictors; only " & ds.n & " are usable. Use forward or stepwise."
            Exit Function
        End If
        For j = 1 To ds.k
            inModel(j) = True
        Next j
        f = FitCurrent(ds, inModel)
        If Not f.ok Then
            msg = "The model with every predictor cannot be estimated: some predictors are constant " & _
                "or perfectly collinear. Remove the redundant columns or use forward or stepwise."
            Exit Function
        End If
        LogStep ds, sl, "Start", "(all candidates)", 0, -1, f
    Else
        f = FitCurrent(ds, inModel)
        LogStep ds, sl, "Start", "(intercept only)", 0, -1, f
    End If

    Do
        changed = False
        If method <> METHOD_BACKWARD Then
            If TryEnter(ds, inModel, crit, enterVal, sl) Then changed = True
        End If
        If method <> METHOD_FORWARD Then
            Do While TryRemove(ds, inModel, crit, removeVal, sl)
                changed = True
                guard = guard + 1
                If guard > MAX_STEPS Then Exit Do
            Loop
        End If
        guard = guard + 1
    Loop While changed And guard <= MAX_STEPS
    sl.hitLimit = (guard > MAX_STEPS)

    RunSelection = True
End Function

Private Function TryEnter(ds As DataSet, inModel() As Boolean, crit As Long, enterVal As Double, _
        sl As StepLog) As Boolean
    Dim j As Long, best As Long, bestT As Double, t As Double, pos As Long
    Dim f As FitResult, bestFit As FitResult, pv As Double

    best = 0
    bestT = -1
    For j = 1 To ds.k
        If Not inModel(j) Then
            inModel(j) = True
            f = FitCurrent(ds, inModel)
            inModel(j) = False
            If f.ok Then
                pos = PositionOf(f, j)
                t = Abs(TStat(f.coef(pos), f.se(pos)))
                ' Candidates share the same residual df, so the largest |t| has the smallest p-value
                If t > bestT Then
                    bestT = t
                    best = j
                    bestFit = f
                End If
            End If
        End If
    Next j
    If best = 0 Then Exit Function

    pos = PositionOf(bestFit, best)
    pv = PValueT(bestFit.coef(pos), bestFit.se(pos), bestFit.dfe)
    If EnterOK(crit, bestT, pv, enterVal) Then
        inModel(best) = True
        LogStep ds, sl, "Entered", ds.xNames(best), TStat(bestFit.coef(pos), bestFit.se(pos)), pv, bestFit
        TryEnter = True
    End If
End Function

Private Function TryRemove(ds As DataSet, inModel() As Boolean, crit As Long, removeVal As Double, _
        sl As StepLog) As Boolean
    Dim f As FitResult, i As Long, worst As Long, worstT As Double, t As Double, pv As Double
    Dim removed As Long, afterFit As FitResult

    f = FitCurrent(ds, inModel)
    If Not f.ok Or f.p = 0 Then Exit Function
    worst = 0
    worstT = 0
    For i = 1 To f.p
        t = Abs(TStat(f.coef(i), f.se(i)))
        If worst = 0 Or t < worstT Then
            worstT = t
            worst = i
        End If
    Next i
    pv = PValueT(f.coef(worst), f.se(worst), f.dfe)
    If RemoveOK(crit, worstT, pv, removeVal) Then
        removed = f.vars(worst)
        inModel(removed) = False
        afterFit = FitCurrent(ds, inModel)
        LogStep ds, sl, "Removed", ds.xNames(removed), TStat(f.coef(worst), f.se(worst)), pv, afterFit
        TryRemove = True
    End If
End Function

Private Function PositionOf(f As FitResult, varIndex As Long) As Long
    Dim i As Long
    For i = 1 To f.p
        If f.vars(i) = varIndex Then
            PositionOf = i
            Exit Function
        End If
    Next i
End Function

Private Sub LogStep(ds As DataSet, sl As StepLog, action As String, varName As String, _
        tv As Double, pv As Double, f As FitResult)
    Dim i As Long
    sl.nSteps = sl.nSteps + 1
    i = sl.nSteps
    ReDim Preserve sl.action(1 To i)
    ReDim Preserve sl.varName(1 To i)
    ReDim Preserve sl.tValue(1 To i)
    ReDim Preserve sl.pValue(1 To i)
    ReDim Preserve sl.nVars(1 To i)
    ReDim Preserve sl.r2(1 To i)
    ReDim Preserve sl.adjR2(1 To i)
    ReDim Preserve sl.s(1 To i)
    ReDim Preserve sl.aic(1 To i)
    ReDim Preserve sl.bic(1 To i)
    sl.action(i) = action
    sl.varName(i) = varName
    sl.tValue(i) = tv
    sl.pValue(i) = pv
    sl.nVars(i) = f.p
    sl.r2(i) = 1 - f.sse / ds.syy
    sl.adjR2(i) = 1 - (f.sse / f.dfe) / (ds.syy / (ds.n - 1))
    sl.s(i) = Sqr(f.sse / f.dfe)
    If f.sse > 0 Then
        sl.aic(i) = ds.n * Log(f.sse / ds.n) + 2 * (f.p + 1)
        sl.bic(i) = ds.n * Log(f.sse / ds.n) + (f.p + 1) * Log(ds.n)
    End If
End Sub

'------------------------------------------------------------------------------
' Least squares
'------------------------------------------------------------------------------
Private Function FitCurrent(ds As DataSet, inModel() As Boolean) As FitResult
    Dim vars() As Long, p As Long, j As Long
    ReDim vars(1 To ds.k)
    For j = 1 To ds.k
        If inModel(j) Then
            p = p + 1
            vars(p) = j
        End If
    Next j
    FitCurrent = FitModel(ds, vars, p)
End Function

Private Function FitModel(ds As DataSet, vars() As Long, p As Long) As FitResult
    Dim f As FitResult, a() As Double, inv() As Double
    Dim i As Long, j As Long, b As Double, mse As Double, v0 As Double, e As Double

    f.p = p
    f.dfe = ds.n - p - 1
    ReDim f.coef(0 To p)
    ReDim f.se(0 To p)
    If f.dfe < 1 Then
        FitModel = f
        Exit Function
    End If

    If p = 0 Then
        f.sse = ds.syy
        f.coef(0) = ds.yBar
        f.se(0) = Sqr(f.sse / f.dfe / ds.n)
        f.ok = True
        FitModel = f
        Exit Function
    End If

    ReDim f.vars(1 To p)
    ReDim a(1 To p, 1 To p)
    For i = 1 To p
        f.vars(i) = vars(i)
        For j = 1 To p
            a(i, j) = ds.cxx(vars(i), vars(j))
        Next j
    Next i
    If Not InvertSPD(a, p, inv) Then
        FitModel = f
        Exit Function
    End If

    For i = 1 To p
        b = 0
        For j = 1 To p
            b = b + inv(i, j) * ds.cxy(vars(j))
        Next j
        f.coef(i) = b
    Next i
    f.coef(0) = ds.yBar
    For i = 1 To p
        f.coef(0) = f.coef(0) - f.coef(i) * ds.xBar(vars(i))
    Next i

    ' Residual sum of squares from the residuals themselves
    f.sse = 0
    For i = 1 To ds.n
        e = ds.y(i) - f.coef(0)
        For j = 1 To p
            e = e - f.coef(j) * ds.x(i, vars(j))
        Next j
        f.sse = f.sse + e * e
    Next i

    mse = f.sse / f.dfe
    v0 = 1 / ds.n
    For i = 1 To p
        For j = 1 To p
            v0 = v0 + ds.xBar(vars(i)) * inv(i, j) * ds.xBar(vars(j))
        Next j
    Next i
    f.se(0) = Sqr(mse * v0)
    For i = 1 To p
        f.se(i) = Sqr(mse * inv(i, i))
    Next i
    f.ok = True
    FitModel = f
End Function

' Gauss-Jordan inverse of a symmetric positive semi-definite matrix, no pivoting.
' Each pivot equals 1 - R^2 of that column on the earlier ones times its diagonal,
' so the tolerance check rejects constant and collinear predictors.
Private Function InvertSPD(a() As Double, m As Long, inv() As Double) As Boolean
    Dim d() As Double, i As Long, j As Long, c As Long, r As Long
    Dim piv As Double, f As Double

    ReDim inv(1 To m, 1 To m)
    ReDim d(1 To m)
    For i = 1 To m
        d(i) = a(i, i)
        inv(i, i) = 1
    Next i
    For c = 1 To m
        piv = a(c, c)
        If d(c) <= 0 Or piv <= SING_TOL * d(c) Then Exit Function
        For j = 1 To m
            a(c, j) = a(c, j) / piv
            inv(c, j) = inv(c, j) / piv
        Next j
        For r = 1 To m
            If r <> c Then
                f = a(r, c)
                If f <> 0 Then
                    For j = 1 To m
                        a(r, j) = a(r, j) - f * a(c, j)
                        inv(r, j) = inv(r, j) - f * inv(c, j)
                    Next j
                End If
            End If
        Next r
    Next c
    InvertSPD = True
End Function

Private Function TStat(coef As Double, se As Double) As Double
    On Error GoTo Huge
    If se > 0 Then
        TStat = coef / se
    ElseIf coef = 0 Then
        TStat = 0
    Else
        GoTo Huge
    End If
    Exit Function
Huge:
    TStat = IIf(coef < 0, -1E+300, 1E+300)
End Function

Private Function PValueT(coef As Double, se As Double, df As Long) As Double
    On Error GoTo Tiny
    PValueT = Application.WorksheetFunction.TDist(Abs(TStat(coef, se)), df, 2)
    Exit Function
Tiny:
    PValueT = 0
End Function

'------------------------------------------------------------------------------
' Report
'------------------------------------------------------------------------------
Private Sub WriteReport(ds As DataSet, method As Long, crit As Long, enterVal As Double, removeVal As Double, _
        inModel() As Boolean, sl As StepLog, dropped As Long, yRng As Range, xRng As Range)
    Dim wb As Workbook, ws As Worksheet, f As FitResult
    Dim r As Long, i As Long, j As Long, tCrit As Double, ssr As Double, top As Long
    Dim fStat As Double, methodName As String, hasExcluded As Boolean
    Dim trial As FitResult, pos As Long

    Select Case method
        Case METHOD_FORWARD: methodName = "Forward selection"
        Case METHOD_BACKWARD: methodName = "Backward elimination"
        Case Else: methodName = "Stepwise"
    End Select

    Set wb = yRng.Worksheet.Parent
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = UniqueSheetName(wb, "Stepwise")
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    f = FitCurrent(ds, inModel)
    ssr = ds.syy - f.sse
    tCrit = Application.WorksheetFunction.TInv(0.05, f.dfe)

    ' Title bar and credit
    With ws.Range("A1:K1")
        .Interior.Color = BURNT_ORANGE
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 14
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 26
    ws.Cells(1, 1).Value = "  Stepwise Regression Output"
    With ws.Cells(2, 1)
        .Value = "  " & CREDIT
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = BURNT_ORANGE
    End With

    r = 4
    PutSection ws, r, "Settings"
    PutPair ws, r, "Dependent variable", ds.yName
    PutPair ws, r, "Method", methodName
    If crit = CRIT_T Then
        PutPair ws, r, "Criterion", "t Stat"
        If method <> METHOD_BACKWARD Then PutPair ws, r, "|t| to enter", enterVal, "0.00"
        If method <> METHOD_FORWARD Then PutPair ws, r, "|t| to remove", removeVal, "0.00"
    Else
        PutPair ws, r, "Criterion", "P-value"
        If method <> METHOD_BACKWARD Then PutPair ws, r, "P-value to enter", enterVal, "0.00"
        If method <> METHOD_FORWARD Then PutPair ws, r, "P-value to remove", removeVal, "0.00"
    End If
    PutPair ws, r, "Observations used", ds.n, "0"
    PutPair ws, r, "Rows excluded (blank or non-numeric)", dropped, "0"
    PutPair ws, r, "Y range", yRng.Address(External:=True)
    PutPair ws, r, "X range", xRng.Address(External:=True)
    If sl.hitLimit Then
        PutPair ws, r, "Warning", "Stopped after " & MAX_STEPS & " iterations; the procedure was cycling."
        ws.Cells(r - 1, 2).Font.Color = RGB(192, 0, 0)
    End If

    ' Step history
    r = r + 1
    PutSection ws, r, "Selection steps"
    PutHeader ws, r, Array("Step", "Action", "Variable", "t Stat", "P-value", "Terms in model", _
        "R Square", "Adjusted R Square", "RMSE", "AIC", "BIC")
    top = r
    For i = 1 To sl.nSteps
        ws.Cells(r, 1).Value = i - 1
        ws.Cells(r, 2).Value = sl.action(i)
        ws.Cells(r, 3).Value = sl.varName(i)
        If sl.pValue(i) >= 0 Then
            ws.Cells(r, 4).Value = sl.tValue(i)
            ws.Cells(r, 5).Value = sl.pValue(i)
        End If
        ws.Cells(r, 6).Value = sl.nVars(i)
        ws.Cells(r, 7).Value = sl.r2(i)
        ws.Cells(r, 8).Value = sl.adjR2(i)
        ws.Cells(r, 9).Value = sl.s(i)
        ws.Cells(r, 10).Value = sl.aic(i)
        ws.Cells(r, 11).Value = sl.bic(i)
        r = r + 1
    Next i
    SetFormats ws, top, r - 1, Array("@", "@", "@", FMT_T, FMT_P, "0", FMT_R2, FMT_R2, FMT_NUM, FMT_NUM, FMT_NUM)

    ' Final model summary
    r = r + 1
    PutSection ws, r, "Final model"
    PutPair ws, r, "Multiple R", Sqr(Application.WorksheetFunction.Max(0, 1 - f.sse / ds.syy)), FMT_R2
    PutPair ws, r, "R Square", 1 - f.sse / ds.syy, FMT_R2
    PutPair ws, r, "Adjusted R Square", 1 - (f.sse / f.dfe) / (ds.syy / (ds.n - 1)), FMT_R2
    PutPair ws, r, "RMSE", Sqr(f.sse / f.dfe), FMT_NUM
    PutPair ws, r, "Observations", ds.n, "0"

    ' ANOVA
    r = r + 1
    PutHeader ws, r, Array("ANOVA", "df", "SS", "MS", "F", "Significance F")
    top = r
    ws.Cells(r, 1).Value = "Regression"
    ws.Cells(r, 2).Value = f.p
    ws.Cells(r, 3).Value = ssr
    If f.p > 0 Then
        ws.Cells(r, 4).Value = ssr / f.p
        If f.sse > 0 Then
            fStat = (ssr / f.p) / (f.sse / f.dfe)
            ws.Cells(r, 5).Value = fStat
            ws.Cells(r, 6).Value = Application.WorksheetFunction.FDist(fStat, f.p, f.dfe)
        End If
    End If
    r = r + 1
    ws.Cells(r, 1).Value = "Residual"
    ws.Cells(r, 2).Value = f.dfe
    ws.Cells(r, 3).Value = f.sse
    ws.Cells(r, 4).Value = f.sse / f.dfe
    r = r + 1
    ws.Cells(r, 1).Value = "Total"
    ws.Cells(r, 2).Value = ds.n - 1
    ws.Cells(r, 3).Value = ds.syy
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Borders(xlEdgeTop).Color = RGB(191, 191, 191)
    SetFormats ws, top, r, Array("@", "0", FMT_NUM, FMT_NUM, FMT_NUM, FMT_P)
    r = r + 2

    ' Coefficients
    PutHeader ws, r, Array("Term", "Coefficients", "Standard Error", "t Stat", "P-value", _
        "Lower 95%", "Upper 95%")
    top = r
    For i = 0 To f.p
        If i = 0 Then ws.Cells(r, 1).Value = "Intercept" Else ws.Cells(r, 1).Value = ds.xNames(f.vars(i))
        ws.Cells(r, 2).Value = f.coef(i)
        ws.Cells(r, 3).Value = f.se(i)
        ws.Cells(r, 4).Value = TStat(f.coef(i), f.se(i))
        ws.Cells(r, 5).Value = PValueT(f.coef(i), f.se(i), f.dfe)
        ws.Cells(r, 6).Value = f.coef(i) - tCrit * f.se(i)
        ws.Cells(r, 7).Value = f.coef(i) + tCrit * f.se(i)
        r = r + 1
    Next i
    SetFormats ws, top, r - 1, Array("@", FMT_COEF, FMT_COEF, FMT_T, FMT_P, FMT_COEF, FMT_COEF)

    ' Variables left out
    For j = 1 To ds.k
        If Not inModel(j) Then hasExcluded = True
    Next j
    If hasExcluded Then
        r = r + 1
        PutSection ws, r, "Variables not in the final model"
        PutHeader ws, r, Array("Variable", "t Stat if added", "P-value if added", "Note")
        top = r
        For j = 1 To ds.k
            If Not inModel(j) Then
                ws.Cells(r, 1).Value = ds.xNames(j)
                inModel(j) = True
                trial = FitCurrent(ds, inModel)
                inModel(j) = False
                If trial.ok Then
                    pos = PositionOf(trial, j)
                    ws.Cells(r, 2).Value = TStat(trial.coef(pos), trial.se(pos))
                    ws.Cells(r, 3).Value = PValueT(trial.coef(pos), trial.se(pos), trial.dfe)
                ElseIf ds.cxx(j, j) <= 0 Then
                    ws.Cells(r, 4).Value = "Constant column"
                Else
                    ws.Cells(r, 4).Value = "Collinear with terms in the model or too few rows"
                End If
                r = r + 1
            End If
        Next j
        SetFormats ws, top, r - 1, Array("@", FMT_T, FMT_P, "@")
    End If

    ' Footer
    r = r + 1
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 11))
        .Borders(xlEdgeTop).Color = BURNT_ORANGE
        .Borders(xlEdgeTop).Weight = xlThin
    End With
    With ws.Cells(r, 1)
        .Value = "Stepwise Regression add-in  |  " & CREDIT
        .Font.Italic = True
        .Font.Size = 8
        .Font.Color = GREY_TEXT
    End With

    ' Fixed widths so the long range addresses overflow instead of widening column B
    ws.Columns("A").ColumnWidth = 32
    ws.Columns("B").ColumnWidth = 13
    ws.Columns("D:K").ColumnWidth = 12
    ws.Columns("C").AutoFit
    If ws.Columns("C").ColumnWidth < 16 Then ws.Columns("C").ColumnWidth = 16
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
End Sub

Private Sub PutSection(ws As Worksheet, r As Long, title As String)
    With ws.Cells(r, 1)
        .Value = title
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = BURNT_ORANGE
    End With
    r = r + 1
End Sub

Private Sub PutPair(ws As Worksheet, r As Long, label As String, v As Variant, _
        Optional numFmt As String = "")
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 1).Font.Color = GREY_TEXT
    ws.Cells(r, 2).Value = v
    If Len(numFmt) > 0 Then ws.Cells(r, 2).NumberFormat = numFmt
    ws.Cells(r, 2).HorizontalAlignment = xlLeft
    r = r + 1
End Sub

Private Sub PutHeader(ws As Worksheet, r As Long, headers As Variant)
    Dim c As Long, n As Long
    n = UBound(headers) - LBound(headers) + 1
    For c = 1 To n
        ws.Cells(r, c).Value = headers(LBound(headers) + c - 1)
    Next c
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, n))
        .Font.Bold = True
        .Interior.Color = ORANGE_TINT
        .Borders(xlEdgeBottom).Color = BURNT_ORANGE
        .Borders(xlEdgeBottom).Weight = xlMedium
        .WrapText = True
        .VerticalAlignment = xlBottom
    End With
    r = r + 1
End Sub

' Applies one number format per column to rows r1..r2 and aligns the header row
' above to match. "@" marks a column to left-align (text, or the step number).
Private Sub SetFormats(ws As Worksheet, r1 As Long, r2 As Long, fmts As Variant)
    Dim c As Long, col As Long
    For c = LBound(fmts) To UBound(fmts)
        col = c - LBound(fmts) + 1
        If fmts(c) = "@" Then
            ws.Range(ws.Cells(r1 - 1, col), ws.Cells(Application.WorksheetFunction.Max(r1, r2), col)).HorizontalAlignment = xlLeft
        Else
            ws.Cells(r1 - 1, col).HorizontalAlignment = xlRight
            If r2 >= r1 Then ws.Range(ws.Cells(r1, col), ws.Cells(r2, col)).NumberFormat = fmts(c)
        End If
    Next c
End Sub

Private Function UniqueSheetName(wb As Workbook, base As String) As String
    Dim nm As String, i As Long, sh As Object
    nm = base
    i = 1
    Do
        Set sh = Nothing
        On Error Resume Next
        Set sh = wb.Sheets(nm)
        On Error GoTo 0
        If sh Is Nothing Then Exit Do
        i = i + 1
        nm = base & " (" & i & ")"
    Loop
    UniqueSheetName = nm
End Function
