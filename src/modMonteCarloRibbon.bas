Attribute VB_Name = "modMonteCarloRibbon"
' =============================================================================
' modMonteCarloRibbon - the Monte Carlo buttons on the Tools menu.
' Copyright (c) 2026 W Phillips, edgewisedata.com. MIT License.
'
' UI ONLY. Every MsgBox and InputBox in this feature is here; every decision it
' makes is in modMonteCarlo, which can be exercised without a dialog appearing.
'
' Ribbon callbacks must be PUBLIC. A private one fails at load with "callback
' not found" and Excel does not say which one it could not find.
'
' Callback names must also be UNIQUE across the whole VBA project. The ribbon
' finds them by bare name, so a Public procedure of the same name in any other
' module makes it fail with "Cannot run the macro" - which is exactly what
' happened when modMonteCarlo also had a McInstallLibrary.
'
' Control ids in ribbon - Monte Carlo menu items.xml:
'   McGalCont   -> McPickDistribution   (Continuous gallery)
'   McGalDisc   -> McPickDistribution   (Discrete gallery)
'   McC_*       -> McPickCopula         (Insert Monte Carlo Copula menu, including
'                                        McC_ClaytonNeg and McC_FrankNeg)
'   McT_*       -> McPickTimeSeries     (Insert Time Series menu:
'                                        simulators and Fit functions)
'   McX_*       -> McPickDiagnostic     (Insert Statistical Diagnostics menu)
'   McStatsBtn          -> McInsertStats                (Insert Monte Carlo Statistics menu)
'   McStatsDetailBtn    -> McInsertStatsDetail
'   McVarTableBtn       -> McInsertVariablesTable
'   McVarTableDetailBtn -> McInsertVariablesTableDetail
'   McHistBtn           -> McInsertHistogram
'   McRiskBtn           -> McInsertRiskMeasures
'   McChartSCurveBtn  -> McChartSCurve      (Insert Monte Carlo Chart menu)
'   McChartHistBtn    -> McChartHistogram
'   McChartTornadoBtn -> McChartTornado
'   McChartFanBtn     -> McChartFan
'   McLibBtn    -> McInstallLibrary
' =============================================================================
Option Explicit

Private Const DIALOG_TITLE As String = "Monte Carlo"


' --- ribbon callbacks --------------------------------------------------------

' Both galleries share this callback. A gallery's onAction takes the clicked
' item's id and position, not just the control, so the signature differs from
' a button's - get it wrong and the ribbon reports it cannot run the macro.
Public Sub McPickDistribution(control As IRibbonControl, id As String, index As Integer)
    Dim idx As Long
    idx = modMonteCarlo.McFindByItemId(id)
    If idx = 0 Then
        Warn "The ribbon item " & id & " is not in the distribution catalog." & vbCrLf & vbCrLf & _
             "The item ids in the ribbon XML and in modMonteCarlo.McCatalog have drifted apart."
        Exit Sub
    End If
    InsertDistribution modMonteCarlo.McCatalogItem(idx), False
End Sub

' Every copula button shares this callback and is told apart by its id.
Public Sub McPickCopula(control As IRibbonControl)
    Dim entry As Variant
    entry = modMonteCarlo.McCopulaByItemId(control.id)
    If IsEmpty(entry) Then
        Warn "The ribbon item " & control.id & " is not in the copula catalog." & vbCrLf & vbCrLf & _
             "The button ids in the ribbon XML and in modMonteCarlo.McCopulaCatalog have drifted apart."
        Exit Sub
    End If
    InsertDistribution entry, True
End Sub

' Every time-series button shares this callback. Simulators go through the
' distribution worker (they take the same Trials / VarID tail); Fit functions
' write a labeled block of estimates from a history the user selects.
Public Sub McPickTimeSeries(control As IRibbonControl)
    Dim entry As Variant
    entry = modMonteCarlo.McTSByItemId(control.id)
    If IsEmpty(entry) Then
        Warn "The ribbon item " & control.id & " is not in the time-series catalog." & vbCrLf & vbCrLf & _
             "The button ids in the ribbon XML and in modMonteCarlo.McTSCatalog have drifted apart."
        Exit Sub
    End If
    If CStr(entry(4)) = "sim" Then
        InsertDistribution entry, False, True
    Else
        InsertTSFit entry
    End If
End Sub

' The six Diagnostics buttons. Data is either one table of history (one variable
' per column; names from the row above) or the spilled results of several
' variables, Ctrl-clicked - the same choice the Variables Table offers.
Public Sub McPickDiagnostic(control As IRibbonControl)
    Dim wb As Workbook
    Dim picked As Range, area As Range, anchor As Range, target As Range
    Dim refs As New Collection, labels As New Collection, seen As New Collection
    Dim dataRef As String, namesRef As String, formulaText As String, key As String, startAt As String
    Dim opt1 As String, opt2 As String, opt3 As String, answer As String, title As String
    Dim n As Long, k As Long, added As Long
    Dim plain As Boolean
    Dim size As Variant

    Dim fromVer As String, why As String
    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub
    title = Mid$(control.id, 5)

    If TypeName(Selection) = "Range" Then startAt = Selection.Address(External:=False)
    On Error Resume Next
    Set picked = Application.InputBox( _
        prompt:="Select the DATA - one variable per column:" & vbCrLf & vbCrLf & _
                "  - a table of history (its header row, if any, names the variables), or" & vbCrLf & _
                "  - any cell of each variable's spilled trials; hold Ctrl to pick several.", _
        title:=DIALOG_TITLE, Default:=startAt, Type:=8)
    On Error GoTo Oops
    If picked Is Nothing Then Exit Sub

    ' One plain block (not part of a spilled result) is a table of history.
    Set anchor = modMonteCarlo.McSpillAnchor(picked.Areas(1))
    plain = (picked.Areas.count = 1 And picked.Cells.count > 1 And Not anchor.HasSpill)
    If plain Then
        n = picked.rows.count
        k = picked.Columns.count
    Else
        For Each area In picked.Areas
            Set anchor = modMonteCarlo.McSpillAnchor(area)
            key = anchor.Worksheet.name & "!" & anchor.Address
            If Not InCollection(seen, key) Then
                seen.Add key, key
                If anchor.HasSpill Then
                    k = k + anchor.SpillingToRange.Columns.count
                    If n = 0 Then n = anchor.SpillingToRange.rows.count
                Else
                    k = k + 1
                End If
            End If
        Next area
    End If
    If n < 3 Then
        Warn "Select at least three rows of data."
        Exit Sub
    End If
    If k < 2 And control.id <> "McX_Normal" Then
        Warn "This diagnostic compares variables: select at least two columns."
        Exit Sub
    End If

    Select Case control.id
        Case "McX_Correl"
            answer = AskText("Pearson or Spearman (rank) correlation?", "Pearson")
            If Len(answer) = 0 Then Exit Sub
            If UCase$(answer) = "SPEARMAN" Then opt2 = """Spearman"""
            answer = AskText("What should the LOWER triangle show? (the upper always shows R)" & vbCrLf & vbCrLf & _
                             "none = R again (a symmetric matrix)" & vbCrLf & _
                             "0 = only the strong pairs, |R| > 0.90" & vbCrLf & _
                             "1 = R squared" & vbCrLf & _
                             "2 = p-value (is the correlation real?)" & vbCrLf & _
                             "3 = Yes / No - significant at alpha?", "none")
            If Len(answer) = 0 Then Exit Sub
            If answer <> "none" Then
                If Not (answer = "0" Or answer = "1" Or answer = "2" Or answer = "3") Then
                    Warn "Type none, 0, 1, 2 or 3."
                    Exit Sub
                End If
                opt1 = answer
                If answer = "3" Then
                    answer = AskText("Significance level alpha?", "0.05")
                    If Len(answer) = 0 Then Exit Sub
                    If answer <> "0.05" Then opt3 = answer
                End If
            End If
        Case "McX_Cov"
            Select Case MsgBox("Sample covariance (divides by n - 1)?" & vbCrLf & vbCrLf & _
                               "Yes: sample (the usual choice).   No: population (divides by n).", _
                               vbYesNoCancel + vbQuestion, DIALOG_TITLE)
                Case vbCancel: Exit Sub
                Case vbNo: opt1 = "FALSE"
            End Select
        Case "McX_Normal"
            answer = AskText("Which tests? SW = Shapiro-Wilk, AD = Anderson-Darling, JB = Jarque-Bera.", "SW,AD,JB")
            If Len(answer) = 0 Then Exit Sub
            If UCase$(Replace(answer, " ", "")) <> "SW,AD,JB" Then opt1 = """" & answer & """"
            If InStr(1, answer, "AD", vbTextCompare) > 0 Then
                answer = AskText("Anderson-Darling variant?" & vbCrLf & vbCrLf & _
                                 "Normal = mean and sd estimated (usual)" & vbCrLf & _
                                 "Unmodified = the same test, reporting the plain A2" & vbCrLf & _
                                 "LogNormal = test the logarithms (positive, right-skewed data)" & vbCrLf & _
                                 "Generic = the data are already standardized (e.g. residuals)", "Normal")
                If Len(answer) = 0 Then Exit Sub
                If UCase$(answer) <> "NORMAL" Then opt2 = """" & answer & """"
            End If
            answer = AskText("Significance level alpha?", "0.05")
            If Len(answer) = 0 Then Exit Sub
            If answer <> "0.05" Then opt3 = answer
        Case "McX_Collin"
            answer = AskText("Which measure? VIF, Tolerance (1 / VIF), R2, or All three.", "VIF")
            If Len(answer) = 0 Then Exit Sub
            If UCase$(answer) <> "VIF" Then opt1 = """" & answer & """"
        Case "McX_Outliers"
            If n > 5000 Then
                If MsgBox("That is " & Format$(n, "#,##0") & " rows - the result has one row per observation. " & _
                          "Continue?", vbOKCancel + vbQuestion, DIALOG_TITLE) = vbCancel Then Exit Sub
            End If
    End Select

    size = modMonteCarlo.McDiagSize(control.id, n, k, Replace(opt1, """", ""))
    Set target = AskTarget("Where should the " & LCase$(title) & " block go?" & vbCrLf & vbCrLf & _
                           BlockSize(CLng(size(0)), CLng(size(1))))
    If target Is Nothing Then Exit Sub
    If Not TargetIsClear(target, CLng(size(0)), CLng(size(1))) Then Exit Sub

    ' References are built only now: whether they need a sheet name depends on the target.
    If plain Then
        dataRef = picked.Address(External:=(picked.Worksheet.name <> target.Worksheet.name))
        If picked.Row > 1 Then
            If Application.WorksheetFunction.CountA(picked.rows(1).offset(-1, 0)) = k Then
                namesRef = picked.rows(1).offset(-1, 0).Address(External:=(picked.Worksheet.name <> target.Worksheet.name))
            End If
        End If
    Else
        Set seen = New Collection
        For Each area In picked.Areas
            Set anchor = modMonteCarlo.McSpillAnchor(area)
            key = anchor.Worksheet.name & "!" & anchor.Address
            If Not InCollection(seen, key) Then
                seen.Add key, key
                refs.Add modMonteCarlo.McSpillReference(anchor, target)
                labels.Add modMonteCarlo.McLabelFragment(anchor, target)
            End If
        Next area
        If refs.count = 1 Then
            dataRef = refs(1)
        Else
            dataRef = "HSTACK(" & JoinRefs(refs) & ")"
            If refs.count = k Then namesRef = "HSTACK(" & JoinRefs(labels) & ")"
        End If
    End If

    AppStateManager.FastModeOn
    If Not modMonteCarlo.McEnsureFunctions(wb, Array(modMonteCarlo.McDiagFunction(control.id)), added, fromVer, why) Then
        AppStateManager.FastModeOff
        Warn "The Monte Carlo functions could not be added to this workbook." & vbCrLf & vbCrLf & why
        Exit Sub
    End If
    formulaText = modMonteCarlo.McBuildDiag(control.id, dataRef, namesRef, opt1, opt2, opt3)
    modMonteCarlo.McWriteFormula target, formulaText
    target.Worksheet.Calculate
    AppStateManager.FastModeOff

    If IsError(target.Value2) Then
        Warn "The result came back as " & target.text & "." & vbCrLf & vbCrLf & _
             "Usual causes: text or blanks in the data, columns of different lengths, a column " & _
             "that never changes, or (for multicollinearity) a variable that is an exact mix of " & _
             "the others. The formula has been left in " & target.Address(False, False) & "."
        Exit Sub
    End If
    MsgBox title & " written to " & target.Address(False, False) & "." & vbCrLf & vbCrLf & _
           formulaText & vbCrLf & vbCrLf & InstallNote(added, fromVer), vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub

Private Function JoinRefs(ByVal c As Collection) As String
    Dim v As Variant
    For Each v In c
        If Len(JoinRefs) > 0 Then JoinRefs = JoinRefs & ", "
        JoinRefs = JoinRefs & CStr(v)
    Next v
End Function

Public Sub McInsertStats(control As IRibbonControl)
    InsertFromSelection "fx.RiskStats", modMonteCarlo.MC_STATS_ROWS, _
        "Select the simulated values first - any cell of the spilled result will do.", True
End Sub

Public Sub McInsertStatsDetail(control As IRibbonControl)
    InsertFromSelection "fx.RiskStatsDetail", modMonteCarlo.MC_STATS_DETAIL_ROWS, _
        "Select the simulated values first - any cell of the spilled result will do.", True
End Sub

Public Sub McInsertVariablesTable(control As IRibbonControl)
    InsertVariablesTable "fx.RiskVariablesTable", modMonteCarlo.MC_STATS_ROWS
End Sub

Public Sub McInsertVariablesTableDetail(control As IRibbonControl)
    InsertVariablesTable "fx.RiskVariablesTableDetail", modMonteCarlo.MC_STATS_DETAIL_ROWS
End Sub

Public Sub McInsertHistogram(control As IRibbonControl)
    InsertFromSelection "fx.RiskHist", modMonteCarlo.MC_HIST_ROWS, _
        "Select the simulated values first - any cell of the spilled result will do.", True
End Sub

' --- the three charts --------------------------------------------------------
'
' Each writes its data block(s) with a LAMBDA, then draws a native chart that
' reads them. The data stays live; so does the chart. The drawing itself is in
' modMonteCarloCharts.

Public Sub McChartHistogram(control As IRibbonControl)
    InsertDistributionChart False
End Sub

Public Sub McChartSCurve(control As IRibbonControl)
    InsertDistributionChart True
End Sub

' Fan chart for a block of paths from a time-series simulator. Asks for the
' band (default P10 to P90) and an optional starting value for period 0, writes
' the fx.RiskChartFan block, and draws the chart beside it.
Public Sub McChartFan(control As IRibbonControl)
    Dim wb As Workbook, ws As Worksheet
    Dim source As Range, anchor As Range, target As Range
    Dim ref As String, outName As String, formulaText As String
    Dim loText As String, hiText As String, startText As String, planText As String, labelText As String
    Dim nPeriods As Long, nRows As Long, added As Long
    Dim lo As Double, hi As Double

    Dim fromVer As String, why As String
    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub
    If Not modXLEdgeHelpers.GetSelectionRange(source, True, DIALOG_TITLE) Then Exit Sub
    If source Is Nothing Then
        Warn "Select the simulated paths first - any cell of the spilled block will do."
        Exit Sub
    End If
    Set anchor = modMonteCarlo.McSpillAnchor(source)
    If anchor.HasSpill Then
        nPeriods = anchor.SpillingToRange.Columns.count
    Else
        nPeriods = source.Columns.count
    End If
    If nPeriods < 2 Then
        Warn "A fan chart needs paths over several periods - a block with one row per trial " & _
             "and one column per period, such as a time-series simulator spills."
        Exit Sub
    End If
    outName = modMonteCarlo.McResultName(anchor)

    loText = AskText("The band's LOWER percentile, between 0 and 1?" & vbCrLf & vbCrLf & _
                     "0.1 draws the band from P10.", "0.1")
    If Len(loText) = 0 Then Exit Sub
    hiText = AskText("The band's UPPER percentile, between 0 and 1?" & vbCrLf & vbCrLf & _
                     "0.9 draws the band up to P90.", "0.9")
    If Len(hiText) = 0 Then Exit Sub
    If Not IsNumeric(loText) Or Not IsNumeric(hiText) Then
        Warn "The percentiles must be numbers between 0 and 1, such as 0.1 and 0.9."
        Exit Sub
    End If
    lo = CDbl(loText)
    hi = CDbl(hiText)
    If lo < 0 Or hi > 1 Or lo >= hi Then
        Warn "The lower percentile must be below the upper one, both between 0 and 1."
        Exit Sub
    End If
    startText = AskText("A starting value to show as period 0, so the fan opens from a point?" & _
                        vbCrLf & vbCrLf & "Type a number or a cell address, or leave none.", "none")
    If Len(startText) = 0 Then Exit Sub
    If LCase$(startText) = "none" Then startText = ""
    planText = AskRangeOrNone("Your single-number forecast - the plan - to draw against the band?" & _
                              vbCrLf & vbCrLf & "Select its " & nPeriods & " values (a row or a column), " & _
                              "or type none.")
    If planText = vbNullChar Then Exit Sub
    labelText = AskRangeOrNone("Labels for the periods, such as dates or month names?" & _
                               vbCrLf & vbCrLf & "Select " & nPeriods & " cells, or type none to " & _
                               "number them 1, 2, 3 ...")
    If labelText = vbNullChar Then Exit Sub

    nRows = modMonteCarlo.MC_FAN_HEAD_ROWS + nPeriods + IIf(Len(startText) > 0, 1, 0)
    Set target = AskTarget("Where should the fan data go?" & vbCrLf & vbCrLf & _
                           BlockSize(nRows, modMonteCarlo.MC_FAN_COLS) & _
                           " The title, a header, then one row per period. The chart is placed to the right.")
    If target Is Nothing Then Exit Sub
    If Not TargetIsClear(target, nRows, modMonteCarlo.MC_FAN_COLS) Then Exit Sub
    Set ws = target.Worksheet

    ref = modMonteCarlo.McSpillReference(anchor, target)
    If Len(ref) = 0 Then ref = source.Address(External:=(source.Worksheet.name <> ws.name))

    AppStateManager.FastModeOn
    If Not modMonteCarlo.McEnsureFunctions(wb, Array("fx.RiskChartFan"), added, fromVer, why) Then
        AppStateManager.FastModeOff
        Warn "The Monte Carlo functions could not be added to this workbook." & vbCrLf & vbCrLf & why
        Exit Sub
    End If
    formulaText = "=" & modMonteCarlo.McFunctionName("fx.RiskChartFan") & "(" & ref & ", " & _
                  NumText(lo) & ", " & NumText(hi) & ", " & _
                  modMonteCarlo.McLabelFragment(anchor, target) & ", " & startText & ", " & _
                  planText & ", " & labelText
    ' Drop trailing empty arguments: "(..., $B$1, , , " -> "(..., $B$1".
    Do While Right$(formulaText, 1) = "," Or Right$(formulaText, 1) = " "
        formulaText = Left$(formulaText, Len(formulaText) - 1)
    Loop
    formulaText = formulaText & ")"
    modMonteCarlo.McWriteFormula target, formulaText
    ws.Calculate
    If IsError(target.Value2) Then
        AppStateManager.FastModeOff
        Warn "The fan data came back as " & target.text & "." & vbCrLf & vbCrLf & _
             "Check that the selection is a block of numbers with at least two trials. " & _
             "The formula has been left in " & target.Address(False, False) & " so you can inspect it."
        Exit Sub
    End If
    modMonteCarloCharts.McDrawFan ws, target, Len(planText) > 0, _
                                  target.offset(0, modMonteCarlo.MC_FAN_COLS + 1).Left, target.top
    AppStateManager.FastModeOff

    MsgBox "Fan chart for " & outName & " written, with the chart beside the data." & vbCrLf & vbCrLf & _
           formulaText & vbCrLf & vbCrLf & _
           "To change the band, edit its two percentiles in the formula in " & _
           target.Address(False, False) & " - the band, the headers and the legend follow. " & _
           "Keep the data where it is: deleting it breaks the chart." & vbCrLf & vbCrLf & _
           InstallNote(added, fromVer), _
           vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub

' A range the user selects, as a sheet-qualified address, or "" for none.
' Returns vbNullChar on Cancel. Type:=0 accepts a selection or typed text.
Private Function AskRangeOrNone(ByVal prompt As String) As String
    Dim v As Variant
    On Error Resume Next
    v = Application.InputBox(prompt:=prompt, title:=DIALOG_TITLE, Default:="none", Type:=0)
    On Error GoTo 0
    If VarType(v) = vbBoolean Then
        AskRangeOrNone = vbNullChar
    ElseIf modMonteCarlo.McIsNone(CStr(v)) Then
        AskRangeOrNone = ""
    Else
        AskRangeOrNone = modMonteCarlo.McArgText(CStr(v))
    End If
End Function

' A number as formula text: Str$ always writes a "." decimal point.
Private Function NumText(ByVal v As Double) As String
    NumText = Trim$(Str$(v))
    If Left$(NumText, 1) = "." Then NumText = "0" & NumText
End Function

' Select the output, then Ctrl-click the inputs. Every input's trials must line
' up with the output's, trial for trial - which they do when they come from
' the same model and trial count.
Public Sub McChartTornado(control As IRibbonControl)
    Dim wb As Workbook, ws As Worksheet
    Dim source As Range, picked As Range, area As Range, anchor As Range, outAnchor As Range
    Dim target As Range
    Dim inRefs As New Collection, labels As New Collection, seen As New Collection
    Dim outRef As String, outName As String, formulaText As String, key As String
    Dim nIn As Long, nCols As Long, added As Long
    Dim useRank As Boolean
    Dim answer As VbMsgBoxResult

    Dim fromVer As String, why As String
    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub
    If Not modXLEdgeHelpers.GetSelectionRange(source, True, DIALOG_TITLE) Then Exit Sub
    If source Is Nothing Then
        Warn "Select the OUTPUT's simulated values first - any cell of its spilled result will do."
        Exit Sub
    End If
    Set outAnchor = modMonteCarlo.McSpillAnchor(source)
    outName = modMonteCarlo.McResultName(outAnchor)

    On Error Resume Next
    Set picked = Application.InputBox( _
        prompt:="Output: " & outName & vbCrLf & vbCrLf & _
                "Now select the INPUTS - any cell of each input's spilled trials." & vbCrLf & _
                "Hold Ctrl to pick several.", _
        title:=DIALOG_TITLE, Type:=8)
    On Error GoTo Oops
    If picked Is Nothing Then Exit Sub

    For Each area In picked.Areas
        Set anchor = modMonteCarlo.McSpillAnchor(area)
        key = anchor.Worksheet.name & "!" & anchor.Address
        If anchor.Address(External:=True) <> outAnchor.Address(External:=True) And Not InCollection(seen, key) Then
            seen.Add key, key
            nIn = nIn + 1
        End If
    Next area
    If nIn = 0 Then
        Warn "No inputs were selected, other than the output itself."
        Exit Sub
    End If

    answer = MsgBox("How should the tornado measure each input's effect?" & vbCrLf & vbCrLf & _
                    "Yes:  swing - the output's mean when the input is in its lowest 10% " & _
                    "versus its highest 10%, in the output's own units." & vbCrLf & vbCrLf & _
                    "No:   rank correlation with the output, from -1 to +1.", _
                    vbYesNoCancel + vbQuestion, DIALOG_TITLE)
    If answer = vbCancel Then Exit Sub
    useRank = (answer = vbNo)
    nCols = IIf(useRank, 4, 7)

    Set target = AskTarget("Where should the tornado data go?" & vbCrLf & vbCrLf & _
                           "It is a " & nCols & "-column block, " & _
                           RowsText(modMonteCarlo.MC_CHART_TITLE_ROWS + nIn + 1) & _
                           " deep. The chart is placed to its right.")
    If target Is Nothing Then Exit Sub
    If Not TargetIsClear(target, modMonteCarlo.MC_CHART_TITLE_ROWS + nIn + 1, nCols) Then Exit Sub
    Set ws = target.Worksheet

    ' Refs are built only now, because whether they need a sheet name depends
    ' on where the result lands.
    Set seen = New Collection
    For Each area In picked.Areas
        Set anchor = modMonteCarlo.McSpillAnchor(area)
        key = anchor.Worksheet.name & "!" & anchor.Address
        If anchor.Address(External:=True) <> outAnchor.Address(External:=True) And Not InCollection(seen, key) Then
            seen.Add key, key
            inRefs.Add modMonteCarlo.McSpillReference(anchor, target)
            labels.Add modMonteCarlo.McLabelFragment(anchor, target)
        End If
    Next area
    outRef = modMonteCarlo.McSpillReference(outAnchor, target)

    AppStateManager.FastModeOn
    If Not modMonteCarlo.McEnsureFunctions(wb, Array("fx.RiskTornado"), added, fromVer, why) Then
        AppStateManager.FastModeOff
        Warn "The Monte Carlo functions could not be added to this workbook." & vbCrLf & vbCrLf & why
        Exit Sub
    End If
    formulaText = modMonteCarlo.McBuildTornado(outRef, inRefs, labels, useRank, _
                                               modMonteCarlo.McLabelFragment(outAnchor, target))
    modMonteCarlo.McWriteFormula target, formulaText
    ws.Calculate
    If IsError(target.Value2) Then
        AppStateManager.FastModeOff
        Warn "The tornado data came back as " & target.text & "." & vbCrLf & vbCrLf & _
             "The usual cause is inputs with a different number of trials from the output. " & _
             "The formula has been left in " & target.Address(False, False) & " so you can inspect it."
        Exit Sub
    End If
    modMonteCarloCharts.McDrawTornado ws, target, target.offset(modMonteCarlo.MC_CHART_TITLE_ROWS, 0), _
        nIn, IIf(useRank, "rank", "swing"), target.offset(0, nCols + 1).Left, target.top
    AppStateManager.FastModeOff

    MsgBox "Tornado for " & outName & " written to " & target.Address(False, False) & _
           ", with the chart beside it." & vbCrLf & vbCrLf & formulaText & vbCrLf & vbCrLf & _
           "The chart reads the block, so it redraws whenever the model changes. " & _
           "Rename an input's header cell and its bar is renamed too; the title " & _
           "follows the output's header cell." & vbCrLf & vbCrLf & _
           InstallNote(added, fromVer), _
           vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub

' VaR, CVaR and ES as one labeled block. Asks for the confidence level and
' for which way round the trials are, since both change the answer.
Public Sub McInsertRiskMeasures(control As IRibbonControl)
    Dim wb As Workbook
    Dim source As Range, target As Range
    Dim ref As String, formulaText As String, answer As String
    Dim conf As Double
    Dim lossesPositive As Boolean
    Dim added As Long
    Dim sign As VbMsgBoxResult

    Dim fromVer As String, why As String
    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub

    If Not modXLEdgeHelpers.GetSelectionRange(source, True, DIALOG_TITLE) Then Exit Sub
    If source Is Nothing Then
        Warn "Select the simulated values first - any cell of the spilled result will do."
        Exit Sub
    End If

    answer = AskText("Confidence level, between 0 and 1?", "0.95")
    If Len(answer) = 0 Then Exit Sub
    If Not IsNumeric(answer) Then
        Warn "The confidence level must be a number between 0 and 1, such as 0.95."
        Exit Sub
    End If
    conf = CDbl(answer)
    If conf <= 0 Or conf >= 1 Then
        Warn "The confidence level must be between 0 and 1, such as 0.95 or 0.99."
        Exit Sub
    End If

    sign = MsgBox("Are these trials profit and loss - gains positive, losses negative?" & vbCrLf & vbCrLf & _
                  "Yes:  P&L. The loss tail is the low end." & vbCrLf & _
                  "No:   loss amounts. The loss tail is the high end." & vbCrLf & vbCrLf & _
                  "Either way the results are reported as positive losses.", _
                  vbYesNoCancel + vbQuestion, DIALOG_TITLE)
    If sign = vbCancel Then Exit Sub
    lossesPositive = (sign = vbNo)

    Set target = AskTarget("Where should the result go?" & vbCrLf & vbCrLf & _
                           BlockSize(modMonteCarlo.MC_RISK_ROWS, 2))
    If target Is Nothing Then Exit Sub
    If Not TargetIsClear(target, modMonteCarlo.MC_RISK_ROWS, 2) Then Exit Sub

    ref = modMonteCarlo.McSpillReference(source, target)
    If Len(ref) = 0 Then
        Warn "Select the simulated values first - any cell of the spilled result will do."
        Exit Sub
    End If

    AppStateManager.FastModeOn
    If Not modMonteCarlo.McEnsureFunctions(wb, Array("fx.RiskMeasures"), added, fromVer, why) Then
        AppStateManager.FastModeOff
        Warn "The Monte Carlo functions could not be added to this workbook." & vbCrLf & vbCrLf & why
        Exit Sub
    End If
    formulaText = modMonteCarlo.McBuildRiskBlock(ref, conf, lossesPositive, _
        modMonteCarlo.McLabelFragment(modMonteCarlo.McSpillAnchor(source), target))
    modMonteCarlo.McWriteFormula target, formulaText
    AppStateManager.FastModeOff

    MsgBox "Written to " & target.Cells(1, 1).Address(False, False) & ":" & vbCrLf & vbCrLf & _
           formulaText & vbCrLf & vbCrLf & InstallNote(added, fromVer), vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub

Public Sub McInstallLibrary(control As IRibbonControl)
    Dim wb As Workbook
    Dim used As Collection
    Dim failed As New Collection, wrapped As New Collection
    Dim added As Long
    Dim msg As String, fromVer As String, why As String
    Dim answer As VbMsgBoxResult

    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub
    If Not modMonteCarlo.McBundleReady(why) Then
        Warn why
        Exit Sub
    End If

    answer = MsgBox("Which Monte Carlo and statistical functions should go into " & wb.name & "?" & vbCrLf & vbCrLf & _
                    "Yes:  only the ones this workbook's formulas use, with what they depend on. " & _
                    "Also repairs a workbook showing #NAME? for a Monte Carlo function." & vbCrLf & vbCrLf & _
                    "No:   the full library (" & modMonteCarlo.McBundleCount() & " functions), " & _
                    "for writing formulas by hand.", _
                    vbYesNoCancel + vbQuestion, DIALOG_TITLE)
    If answer = vbCancel Then Exit Sub

    AppStateManager.FastModeOn
    If answer = vbYes Then
        Set used = modMonteCarlo.McNamesUsedIn(wb)
        If used.count = 0 Then
            AppStateManager.FastModeOff
            MsgBox "No formula in " & wb.name & " uses a Monte Carlo function yet, so there is " & _
                   "nothing to install." & vbCrLf & vbCrLf & _
                   "Inserting a distribution installs what it needs automatically.", _
                   vbInformation, DIALOG_TITLE
            Exit Sub
        End If
        If Not modMonteCarlo.McEnsureFunctions(wb, used, added, fromVer, why) Then
            AppStateManager.FastModeOff
            Warn "The Monte Carlo functions could not be added to this workbook." & vbCrLf & vbCrLf & why
            Exit Sub
        End If
        AppStateManager.FastModeOff
        msg = "This workbook's formulas use " & used.count & " Monte Carlo function" & _
              IIf(used.count = 1, "", "s") & "." & vbCrLf & vbCrLf & _
              IIf(added = 0 And Len(fromVer) = 0, "All of them, and everything they depend on, " & _
                  "were already installed.", InstallNote(added, fromVer))
    Else
        added = modMonteCarlo.McInstallAll(wb, failed, wrapped)
        AppStateManager.FastModeOff
        msg = added & " Monte Carlo functions are now in " & wb.name & "." & vbCrLf & vbCrLf & _
              "They are ordinary defined names, so this workbook will keep calculating " & _
              "on a machine with no add-ins at all."
        If failed.count > 0 Then msg = msg & vbCrLf & vbCrLf & _
              failed.count & " could not be added." & NameList(failed)
        If wrapped.count > 0 Then msg = msg & vbCrLf & vbCrLf & _
              "Stored as zero-argument functions, so call them with brackets:" & NameList(wrapped)
    End If

    MsgBox msg, vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub


' --- developer build steps ---------------------------------------------------
'
' Run from the Immediate window while BUILDING the add-in, not from the ribbon:
'
'   McDevRefreshBundle          load MonteCarlo.txt into the bundle; save after
'   McDevRemoveFromLambdaLibrary   one-off: drop the fx.Risk rows from tblAlonzoChurch
'
' Both change the add-in itself, so both confirm first and neither saves.

Public Sub McDevRefreshBundle(Optional ByVal path As String = "")
    Dim n As Long, why As String
    If Len(path) = 0 Then path = modMonteCarlo.MC_SOURCE_PATH
    If Not modMonteCarlo.McLoadBundleFromFile(path, n, why) Then
        Warn why
        Exit Sub
    End If
    MsgBox n & " Monte Carlo functions, version " & modMonteCarlo.MC_LIBRARY_VERSION & _
           ", are now bundled in " & ThisWorkbook.name & "." & vbCrLf & vbCrLf & _
           "Save the add-in to keep them.", vbInformation, DIALOG_TITLE
End Sub

Public Sub McDevRemoveFromLambdaLibrary()
    Dim removed As Long, why As String
    If Not modXLEdgeHelpers.ConfirmDestructive( _
            "Remove every fx.Risk function from XL Edge's own LAMBDA library (tblAlonzoChurch)?" & _
            vbCrLf & vbCrLf & "The Monte Carlo functions are installed from the bundle now, so " & _
            "the library no longer needs them. Nothing else in it is touched.", DIALOG_TITLE) Then
        Exit Sub
    End If
    If Not modMonteCarlo.McRemoveFromLambdaLibrary(removed, why) Then
        Warn why
        Exit Sub
    End If
    MsgBox IIf(removed = 0, "There were no fx.Risk functions in the LAMBDA library.", _
               removed & " fx.Risk functions were removed from the LAMBDA library.") & _
           vbCrLf & vbCrLf & "Save the add-in to keep the change.", vbInformation, DIALOG_TITLE
End Sub


' --- worker for the two histogram charts -------------------------------------

' Both charts read ONE block, written by fx.RiskChartHist: the title, the
' P10 / P50 / P90 lines and the bins in a single spill, so the lines cannot drift
' from the bars and no second block can be overrun. The chart follows the block
' even if the formula's Bins is changed later (see modMonteCarloCharts).
Private Sub InsertDistributionChart(ByVal isSCurve As Boolean)
    Dim wb As Workbook, ws As Worksheet
    Dim source As Range, target As Range
    Dim ref As String, outName As String
    Dim nRows As Long, nCols As Long, added As Long

    Dim fromVer As String, why As String
    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub
    If Not modXLEdgeHelpers.GetSelectionRange(source, True, DIALOG_TITLE) Then Exit Sub
    If source Is Nothing Then
        Warn "Select the simulated values first - any cell of the spilled result will do."
        Exit Sub
    End If
    outName = modMonteCarlo.McResultName(modMonteCarlo.McSpillAnchor(source))

    nRows = modMonteCarlo.MC_CHART_HEAD_ROWS + modMonteCarlo.MC_CHART_BINS
    nCols = modMonteCarlo.MC_CHART_HIST_COLS
    Set target = AskTarget("Where should the chart data go?" & vbCrLf & vbCrLf & _
        "It is a " & nCols & "-column block, " & nRows & " rows deep: the title, the P10 / P50 / P90 " & _
        "lines, then " & modMonteCarlo.MC_CHART_BINS & " bins. The chart is placed to the right.")
    If target Is Nothing Then Exit Sub
    If Not TargetIsClear(target, nRows, nCols) Then Exit Sub
    Set ws = target.Worksheet

    ref = modMonteCarlo.McSpillReference(source, target)
    If Len(ref) = 0 Then
        Warn "Select the simulated values first - any cell of the spilled result will do."
        Exit Sub
    End If

    AppStateManager.FastModeOn
    If Not modMonteCarlo.McEnsureFunctions(wb, Array("fx.RiskChartHist"), added, fromVer, why) Then
        AppStateManager.FastModeOff
        Warn "The Monte Carlo functions could not be added to this workbook." & vbCrLf & vbCrLf & why
        Exit Sub
    End If

    ' The fifth argument names the result: the chart title is linked to it.
    modMonteCarlo.McWriteFormula target, "=" & modMonteCarlo.McFunctionName("fx.RiskChartHist") & _
        "(" & ref & ", , , , " & modMonteCarlo.McLabelFragment(modMonteCarlo.McSpillAnchor(source), target) & ")"
    ws.Calculate

    If IsError(target.Value2) Then
        AppStateManager.FastModeOff
        Warn "The chart data came back as an error (" & target.text & ")." & vbCrLf & vbCrLf & _
             "Check that the selection holds numbers, and that they are not all the same value."
        Exit Sub
    End If

    If isSCurve Then
        modMonteCarloCharts.McDrawHistogramSCurve ws, target, target.offset(0, nCols + 1).Left, target.top
    Else
        modMonteCarloCharts.McDrawOutcomeHistogram ws, target, modMonteCarlo.MC_CHART_MARKS, _
            target.offset(0, nCols + 1).Left, target.top
    End If
    AppStateManager.FastModeOff

    MsgBox IIf(isSCurve, "Histogram and S-curve", "Outcome histogram") & " for " & outName & _
           " written, with the chart beside the data." & vbCrLf & vbCrLf & _
           "The chart reads the spilled data, so it redraws whenever the model changes. " & _
           "To change the number of bins, edit the second argument of the formula in " & _
           target.Address(False, False) & " - the chart follows; leave room below for the extra rows. " & _
           "Keep the data where it is: deleting it breaks the chart." & vbCrLf & vbCrLf & _
           InstallNote(added, fromVer), _
           vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub

Private Function InCollection(ByVal c As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = c(key)
    InCollection = (Err.Number = 0)
    Err.Clear
End Function


' --- worker for the distribution galleries and the copula menu --------------

' entry is a catalog row: Array(display name, base name, params, item id).
' A copula differs only in its result - several columns of uniforms rather
' than one column of values - so only the size message and closing hint change.
' A time-series simulator likewise: one row per trial, one column per period.
Private Sub InsertDistribution(ByVal entry As Variant, ByVal isCopula As Boolean, _
                               Optional ByVal isTimeSeries As Boolean = False)
    Dim wb As Workbook
    Dim i As Long, vid As Long, added As Long, trials As Long, nCols As Long, nTrailing As Long
    Dim params() As String
    Dim args() As String
    Dim one As String, trialsExpr As String, formulaText As String
    Dim target As Range

    Dim fromVer As String, why As String
    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub

    params = Split(CStr(entry(2)), "|")
    ReDim args(LBound(params) To UBound(params))
    For i = LBound(params) To UBound(params)
        one = AskArgument(CStr(entry(0)), params(i), i + 1, UBound(params) + 1)
        If Len(one) = 0 Then Exit Sub
        ' An empty argument - "fx.RiskBeta<lambda>(2, 3, , , ...)" - is how a LAMBDA
        ' sees an optional parameter as omitted.
        If modMonteCarlo.McParamIsOptional(params(i)) And modMonteCarlo.McIsNone(one) Then one = ""
        args(i) = one
        ' Trailing parameters (a copula's Rotation) are written after the tail.
        If modMonteCarlo.McParamIsTrailing(params(i)) Then nTrailing = nTrailing + 1
    Next i

    If modMonteCarlo.McTrialsDefined(wb) Then
        trialsExpr = modMonteCarlo.MC_TRIALS_NAME
    Else
        trials = AskTrials()
        If trials < 1 Then Exit Sub
        trialsExpr = modMonteCarlo.McDefineTrials(wb, trials)
    End If

    trials = modMonteCarlo.McTrialsCount(wb)
    If isTimeSeries Then
        nCols = modMonteCarlo.McTSColumns(params, args)
        If nCols > 0 Then
            Set target = AskTarget("Where should the paths go?" & vbCrLf & vbCrLf & _
                                   "One row per trial, one column per period. " & BlockSize(trials, nCols))
        Else
            Set target = AskTarget("Where should the paths go?" & vbCrLf & vbCrLf & _
                                   "One row per trial (" & RowsText(trials) & "), one column per period.")
        End If
    ElseIf isCopula Then
        nCols = modMonteCarlo.McCopulaColumns(CStr(entry(1)), args)
        If nCols > 0 Then
            Set target = AskTarget("Where should the correlated uniforms go?" & vbCrLf & vbCrLf & _
                                   BlockSize(trials, nCols))
        Else
            Set target = AskTarget("Where should the correlated uniforms go?" & vbCrLf & vbCrLf & _
                                   "It spills one column per variable, " & RowsText(trials) & " deep.")
        End If
    Else
        nCols = 1
        Set target = AskTarget("Where should the trials go?" & vbCrLf & vbCrLf & _
                               "They spill DOWN from the cell you pick, " & RowsText(trials) & _
                               ", so leave room below it.")
    End If
    If target Is Nothing Then Exit Sub
    If Not TargetIsClear(target, trials, nCols) Then Exit Sub

    AppStateManager.FastModeOn

    If Not modMonteCarlo.McEnsureFunctions(wb, Array(CStr(entry(1))), added, fromVer, why) Then
        AppStateManager.FastModeOff
        Warn "The Monte Carlo functions could not be added to this workbook, so " & _
             "the formula was not written." & vbCrLf & vbCrLf & why
        Exit Sub
    End If

    vid = modMonteCarlo.McNextVarId(wb)
    formulaText = modMonteCarlo.McBuildCall(CStr(entry(1)), args, trialsExpr, vid, nTrailing)
    modMonteCarlo.McWriteFormula target, formulaText

    AppStateManager.FastModeOff

    MsgBox CStr(entry(0)) & " written to " & target.Cells(1, 1).Address(False, False) & _
           " as VarID " & vid & "." & vbCrLf & vbCrLf & formulaText & vbCrLf & vbCrLf & _
           InstallNote(added, fromVer) & _
           "Trial count comes from " & modMonteCarlo.MC_TRIALS_NAME & _
           ", which you can change in Name Manager." & CopulaHint(isCopula, target) & _
           TimeSeriesHint(isTimeSeries, target), _
           vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub


' --- shared worker for the result helpers -----------------------------------

' nRows is how deep the two-column block will be. The first dialog states it,
' so the user can pick a spot with room and not land on #SPILL!. withName adds
' the result's header cell as the third argument (fx.RiskStats and
' fx.RiskStatsDetail), so the block's name stays live and the LAMBDA never has
' to look above the trials itself.
Private Sub InsertFromSelection(ByVal baseName As String, ByVal nRows As Long, ByVal hint As String, _
                                Optional ByVal withName As Boolean = False)
    Dim wb As Workbook
    Dim source As Range, target As Range
    Dim ref As String, formulaText As String
    Dim added As Long

    Dim fromVer As String, why As String
    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub

    If Not modXLEdgeHelpers.GetSelectionRange(source, True, DIALOG_TITLE) Then Exit Sub
    If source Is Nothing Then
        Warn hint
        Exit Sub
    End If

    Set target = AskTarget("Where should the result go?" & vbCrLf & vbCrLf & BlockSize(nRows, 2))
    If target Is Nothing Then Exit Sub
    If Not TargetIsClear(target, nRows, 2) Then Exit Sub

    ref = modMonteCarlo.McSpillReference(source, target)
    If Len(ref) = 0 Then
        Warn hint
        Exit Sub
    End If

    AppStateManager.FastModeOn
    If Not modMonteCarlo.McEnsureFunctions(wb, Array(baseName), added, fromVer, why) Then
        AppStateManager.FastModeOff
        Warn "The Monte Carlo functions could not be added to this workbook." & vbCrLf & vbCrLf & why
        Exit Sub
    End If
    formulaText = "=" & modMonteCarlo.McFunctionName(baseName) & "(" & ref & _
                  IIf(withName, ", , " & modMonteCarlo.McLabelFragment(modMonteCarlo.McSpillAnchor(source), target), "") & ")"
    modMonteCarlo.McWriteFormula target, formulaText
    AppStateManager.FastModeOff

    MsgBox "Written to " & target.Cells(1, 1).Address(False, False) & ":" & vbCrLf & vbCrLf & _
           formulaText & vbCrLf & vbCrLf & InstallNote(added, fromVer), vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub


' --- worker for the two variables tables -------------------------------------

' Ctrl-click any cell of each variable's spilled trials; one column per
' variable, in the order picked, under a single label column. The current
' selection is offered as the starting answer. nRows is the block's depth.
Private Sub InsertVariablesTable(ByVal baseName As String, ByVal nRows As Long)
    Dim wb As Workbook
    Dim picked As Range, area As Range, anchor As Range, target As Range
    Dim refs As New Collection, labels As New Collection, seen As New Collection
    Dim formulaText As String, key As String, startAt As String
    Dim nVars As Long, added As Long

    Dim fromVer As String, why As String
    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub

    If TypeName(Selection) = "Range" Then startAt = Selection.Address(External:=False)
    On Error Resume Next
    Set picked = Application.InputBox( _
        prompt:="Select the VARIABLES - any cell of each variable's spilled trials." & vbCrLf & _
                "Hold Ctrl to pick several. Each becomes a column, in the order picked." & vbCrLf & vbCrLf & _
                "Every variable needs the same number of trials.", _
        title:=DIALOG_TITLE, Default:=startAt, Type:=8)
    On Error GoTo Oops
    If picked Is Nothing Then Exit Sub

    For Each area In picked.Areas
        Set anchor = modMonteCarlo.McSpillAnchor(area)
        key = anchor.Worksheet.name & "!" & anchor.Address
        If Not InCollection(seen, key) Then
            seen.Add key, key
            nVars = nVars + 1
        End If
    Next area

    Set target = AskTarget("Where should the table go?" & vbCrLf & vbCrLf & BlockSize(nRows, nVars + 1) & _
                           vbCrLf & "The first column holds the labels, then one column per variable.")
    If target Is Nothing Then Exit Sub
    If Not TargetIsClear(target, nRows, nVars + 1) Then Exit Sub

    ' Refs are built only now, because whether they need a sheet name depends
    ' on where the result lands.
    Set seen = New Collection
    For Each area In picked.Areas
        Set anchor = modMonteCarlo.McSpillAnchor(area)
        key = anchor.Worksheet.name & "!" & anchor.Address
        If Not InCollection(seen, key) Then
            seen.Add key, key
            refs.Add modMonteCarlo.McSpillReference(anchor, target)
            labels.Add modMonteCarlo.McLabelFragment(anchor, target)
        End If
    Next area

    AppStateManager.FastModeOn
    If Not modMonteCarlo.McEnsureFunctions(wb, Array(baseName), added, fromVer, why) Then
        AppStateManager.FastModeOff
        Warn "The Monte Carlo functions could not be added to this workbook." & vbCrLf & vbCrLf & why
        Exit Sub
    End If
    formulaText = modMonteCarlo.McBuildVariablesTable(baseName, refs, labels)
    modMonteCarlo.McWriteFormula target, formulaText
    AppStateManager.FastModeOff

    MsgBox nVars & IIf(nVars = 1, " variable", " variables") & " written to " & _
           target.Cells(1, 1).Address(False, False) & ":" & vbCrLf & vbCrLf & _
           formulaText & vbCrLf & vbCrLf & _
           "Rename a variable's header cell and its column heading follows." & vbCrLf & vbCrLf & _
           InstallNote(added, fromVer), vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub


' --- prompts -----------------------------------------------------------------

' Type:=0 lets the user either type a number or click a cell, and hands back a
' formula fragment either way.
Private Function AskArgument(ByVal distName As String, ByVal paramName As String, _
                             ByVal n As Long, ByVal total As Long) As String
    Dim v As Variant

    On Error Resume Next
    v = Application.InputBox( _
            prompt:=distName & vbCrLf & vbCrLf & _
                    modMonteCarlo.McParamLabel(paramName) & "   (" & n & " of " & total & ")" & vbCrLf & vbCrLf & _
                    "Type a value, or click a cell to reference it." & _
                    IIf(modMonteCarlo.McParamIsOptional(paramName), _
                        vbCrLf & "Optional: type none to leave it out.", ""), _
            title:=DIALOG_TITLE, Type:=0)
    On Error GoTo 0

    If VarType(v) = vbBoolean Then Exit Function        ' Cancel
    AskArgument = modMonteCarlo.McArgText(CStr(v))
End Function

Private Function AskTrials() As Long
    Dim answer As String

    answer = AskText("How many trials?" & vbCrLf & vbCrLf & _
                     "This is stored as " & modMonteCarlo.MC_TRIALS_NAME & _
                     " and drives every distribution in the workbook, so you " & _
                     "only answer this once.", CStr(modMonteCarlo.MC_DEF_TRIALS))
    If Len(answer) = 0 Then Exit Function
    If Not IsNumeric(answer) Then
        Warn "The trial count must be a whole number."
        Exit Function
    End If
    If CDbl(answer) < 1 Or CDbl(answer) <> Int(CDbl(answer)) Then
        Warn "The trial count must be a whole number of 1 or more."
        Exit Function
    End If
    AskTrials = CLng(answer)
End Function

Private Function AskTarget(ByVal prompt As String) As Range
    Dim rg As Range
    On Error Resume Next
    Set rg = Application.InputBox(prompt:=prompt, title:=DIALOG_TITLE, Type:=8)
    On Error GoTo 0
    Set AskTarget = rg
End Function

' A spilling formula lands on #SPILL! if anything is in its way, and the cause
' is not obvious from the error. So the WHOLE area the result will occupy is
' checked, not just the cell it is written to. rows or cols of 0 means the
' size is unknown, and only the first cell is checked.
'
' Clearing the area is offered, never assumed: the user confirms first, and
' only then are the cells in the way emptied.
Private Function TargetIsClear(ByVal target As Range, ByVal nRows As Long, ByVal nCols As Long) As Boolean
    Dim area As Range
    Dim ws As Worksheet
    Set ws = target.Worksheet

    If nRows < 1 Or nCols < 1 Then
        Set area = target.Cells(1, 1)
    Else
        If target.Row + nRows - 1 > ws.rows.count Or target.Column + nCols - 1 > ws.Columns.count Then
            Warn "There is not room for " & nRows & " rows by " & nCols & " columns below and to the " & _
                 "right of " & target.Cells(1, 1).Address(False, False) & "." & vbCrLf & vbCrLf & _
                 "Pick a cell higher up or further left."
            Exit Function
        End If
        Set area = target.Cells(1, 1).Resize(nRows, nCols)
    End If

    If Application.WorksheetFunction.CountA(area) = 0 Then
        TargetIsClear = True
        Exit Function
    End If
    If Not modXLEdgeHelpers.ConfirmDestructive( _
            area.Address(False, False) & " is not empty, and the result needs all of it " & _
            "or it will show #SPILL!." & vbCrLf & vbCrLf & _
            "Clear " & area.Address(False, False) & " and write the result there?", DIALOG_TITLE) Then
        Exit Function
    End If

    On Error Resume Next
    area.ClearContents
    On Error GoTo 0
    TargetIsClear = (Application.WorksheetFunction.CountA(area) = 0)
    If Not TargetIsClear Then
        Warn "Some of " & area.Address(False, False) & " could not be cleared - it is probably " & _
             "part of another formula's spilled result. Pick a different cell."
    End If
End Function

' "It is a two-column block, 11 rows deep." - the size of a result, in words.
Private Function BlockSize(ByVal nRows As Long, ByVal nCols As Long) As String
    Dim w As String
    Select Case nCols
        Case 1: w = "one-column"
        Case 2: w = "two-column"
        Case Else: w = nCols & "-column"
    End Select
    BlockSize = "It is a " & w & " block, " & RowsText(nRows) & " deep."
End Function

Private Function RowsText(ByVal nRows As Long) As String
    If nRows > 0 Then
        RowsText = Format$(nRows, "#,##0") & " rows"
    Else
        RowsText = "one row per trial"
    End If
End Function

' After a time series is written: how to read one period, with the real address.
Private Function TimeSeriesHint(ByVal isTimeSeries As Boolean, ByVal target As Range) As String
    If Not isTimeSeries Then Exit Function
    TimeSeriesHint = vbCrLf & vbCrLf & "Each row is one path, each column one period. " & _
                     "To summarize the last period:" & vbCrLf & _
                     "   =" & modMonteCarlo.McFunctionName("fx.RiskStats") & "(TAKE(" & _
                     target.Cells(1, 1).Address(False, False) & "#, , -1))"
End Function

' A Fit function: ask for the history and options, size the estimates block,
' then write it. No trial count or VarID - the result is a handful of numbers.
Private Sub InsertTSFit(ByVal entry As Variant)
    Dim wb As Workbook
    Dim i As Long, nRows As Long, added As Long
    Dim params() As String
    Dim args() As String
    Dim one As String, formulaText As String
    Dim target As Range

    Dim fromVer As String, why As String
    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub

    params = Split(CStr(entry(2)), "|")
    ReDim args(LBound(params) To UBound(params))
    For i = LBound(params) To UBound(params)
        one = AskArgument(CStr(entry(0)), params(i), i + 1, UBound(params) + 1)
        If Len(one) = 0 Then Exit Sub
        If modMonteCarlo.McParamIsOptional(params(i)) And modMonteCarlo.McIsNone(one) Then one = ""
        args(i) = one
    Next i

    nRows = modMonteCarlo.McTSFitRows(entry, args)
    If nRows > 0 Then
        Set target = AskTarget("Where should the estimates go?" & vbCrLf & vbCrLf & _
                               "A Parameter / Estimate block. " & BlockSize(nRows, 2))
    Else
        Set target = AskTarget("Where should the estimates go?" & vbCrLf & vbCrLf & _
                               "A two-column Parameter / Estimate block.")
    End If
    If target Is Nothing Then Exit Sub
    If Not TargetIsClear(target, nRows, 2) Then Exit Sub

    AppStateManager.FastModeOn
    If Not modMonteCarlo.McEnsureFunctions(wb, Array(CStr(entry(1))), added, fromVer, why) Then
        AppStateManager.FastModeOff
        Warn "The Monte Carlo functions could not be added to this workbook, so " & _
             "the formula was not written." & vbCrLf & vbCrLf & why
        Exit Sub
    End If
    formulaText = modMonteCarlo.McBuildFitCall(CStr(entry(1)), args)
    modMonteCarlo.McWriteFormula target, formulaText
    target.Worksheet.Calculate
    AppStateManager.FastModeOff

    If IsError(target.Value2) Then
        Warn "The estimates came back as " & target.text & "." & vbCrLf & vbCrLf & _
             "Usual causes: too short a history, values of 0 or less where prices are " & _
             "needed, or a series that does not behave like the model (a trending series " & _
             "has no mean reversion to estimate). The formula has been left in " & _
             target.Address(False, False) & " so you can inspect it."
        Exit Sub
    End If
    MsgBox CStr(entry(0)) & " written to " & target.Address(False, False) & "." & vbCrLf & vbCrLf & _
           formulaText & vbCrLf & vbCrLf & _
           "Point the matching simulator's arguments at these cells to simulate " & _
           "from the fitted model." & vbCrLf & vbCrLf & InstallNote(added, fromVer), _
           vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub

' After a copula is written: how to use it, with the real cell address.
Private Function CopulaHint(ByVal isCopula As Boolean, ByVal target As Range) As String
    If Not isCopula Then Exit Function
    CopulaHint = vbCrLf & vbCrLf & "Feed one column into each distribution's Trials argument:" & vbCrLf & _
                 "   =" & modMonteCarlo.McFunctionName("fx.RiskNormal") & "(100, 15, CHOOSECOLS(" & _
                 target.Cells(1, 1).Address(False, False) & "#, 1))"
End Function


' --- guards and plumbing -----------------------------------------------------

Private Function HaveWorkbook(ByRef wb As Workbook) As Boolean
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Warn "Open a workbook first."
        Exit Function
    End If
    HaveWorkbook = True
End Function

' InputBox, not Application.InputBox: StrPtr distinguishes Cancel from an
' answer the user deliberately left empty.
Private Function AskText(ByVal prompt As String, ByVal defaultText As String) As String
    Dim answer As String
    answer = InputBox(prompt, DIALOG_TITLE, defaultText)
    If StrPtr(answer) = 0 Then Exit Function
    AskText = Trim$(answer)
End Function

' What the install step did, for the closing message - or "" if nothing.
Private Function InstallNote(ByVal added As Long, ByVal fromVer As String) As String
    If Len(fromVer) > 0 Then
        InstallNote = "This workbook's Monte Carlo functions were updated from " & fromVer & _
                      " to " & modMonteCarlo.MC_LIBRARY_VERSION & ". Check any statistics blocks: " & _
                      "their layout can change between versions."
    ElseIf added > 0 Then
        InstallNote = added & " Monte Carlo function" & IIf(added = 1, " was", "s were") & _
                      " added to this workbook, so it keeps working for people who do not have XL Edge."
    End If
End Function

Private Sub Warn(ByVal msg As String)
    MsgBox msg, vbExclamation, DIALOG_TITLE
End Sub

Private Sub ShowError()
    Dim n As Long, t As String
    n = Err.Number
    t = Err.description
    AppStateManager.FastModeOff
    modXLEdgeHelpers.ReportError n, t, DIALOG_TITLE
End Sub

Private Function NameList(ByVal names As Collection) As String
    Dim s As String, v As Variant
    For Each v In names
        s = s & vbCrLf & "   " & CStr(v)
    Next v
    NameList = s
End Function
