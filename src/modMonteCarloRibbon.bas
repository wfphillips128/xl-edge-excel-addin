Attribute VB_Name = "modMonteCarloRibbon"
' =============================================================================
' modMonteCarloRibbon - the five Monte Carlo buttons on the Tools menu.
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
'   McStatsBtn  -> McInsertStats
'   McHistBtn   -> McInsertHistogram
'   McRiskBtn   -> McInsertRiskMeasures
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
        Warn "The ribbon item " & id & " is not in the distribution catalogue." & vbCrLf & vbCrLf & _
             "The item ids in the ribbon XML and in modMonteCarlo.McCatalog have drifted apart."
        Exit Sub
    End If
    InsertDistribution idx
End Sub

Public Sub McInsertStats(control As IRibbonControl)
    InsertFromSelection "fx.RiskStats", _
        "Select the simulated values first - any cell of the spilled result will do."
End Sub

Public Sub McInsertHistogram(control As IRibbonControl)
    InsertFromSelection "fx.RiskHist", _
        "Select the simulated values first - any cell of the spilled result will do."
End Sub

' VaR, CVaR and ES as one labelled block. Asks for the confidence level and
' for which way round the trials are, since both change the answer.
Public Sub McInsertRiskMeasures(control As IRibbonControl)
    Dim wb As Workbook
    Dim source As Range, target As Range
    Dim ref As String, formulaText As String, answer As String
    Dim conf As Double
    Dim lossesPositive As Boolean
    Dim added As Long
    Dim sign As VbMsgBoxResult

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
                           "It is a two-column block, four rows deep.")
    If target Is Nothing Then Exit Sub
    If Not TargetIsClear(target) Then Exit Sub

    ref = modMonteCarlo.McSpillReference(source, target)
    If Len(ref) = 0 Then
        Warn "Select the simulated values first - any cell of the spilled result will do."
        Exit Sub
    End If

    AppStateManager.FastModeOn
    If Not modMonteCarlo.McEnsureLibrary(wb, added) Then
        AppStateManager.FastModeOff
        Warn "The Monte Carlo functions could not be added to this workbook."
        Exit Sub
    End If
    formulaText = modMonteCarlo.McBuildRiskBlock(ref, conf, lossesPositive)
    modMonteCarlo.McWriteFormula target, formulaText
    AppStateManager.FastModeOff

    MsgBox "Written to " & target.Cells(1, 1).Address(False, False) & ":" & vbCrLf & vbCrLf & _
           formulaText, vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub

Public Sub McInstallLibrary(control As IRibbonControl)
    Dim wb As Workbook
    Dim added As Long
    Dim failed As New Collection
    Dim wrapped As New Collection
    Dim msg As String

    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub

    If modMonteCarlo.McLibraryInstalled(wb) Then
        If Not modXLEdgeHelpers.ConfirmProceed( _
                "The Monte Carlo functions are already in this workbook." & vbCrLf & vbCrLf & _
                "Download the published library again and overwrite them?", DIALOG_TITLE) Then
            Exit Sub
        End If
    End If

    AppStateManager.FastModeOn
    added = modMonteCarlo.McDownloadLibrary(wb, failed, wrapped)
    AppStateManager.FastModeOff

    msg = added & " Monte Carlo functions are now in " & wb.name & "." & vbCrLf & vbCrLf & _
          "They are ordinary defined names, so this workbook will keep calculating " & _
          "on a machine with no add-ins at all."
    If failed.Count > 0 Then msg = msg & vbCrLf & vbCrLf & _
          failed.Count & " could not be added." & NameList(failed)
    If wrapped.Count > 0 Then msg = msg & vbCrLf & vbCrLf & _
          "Stored as zero-argument functions, so call them with brackets:" & NameList(wrapped)

    MsgBox msg, vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub


' --- worker for the distribution galleries -----------------------------------

Private Sub InsertDistribution(ByVal idx As Long)
    Dim wb As Workbook
    Dim i As Long, vid As Long, added As Long, trials As Long
    Dim entry As Variant
    Dim params() As String
    Dim args() As String
    Dim one As String, trialsExpr As String, formulaText As String
    Dim target As Range

    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub

    entry = modMonteCarlo.McCatalogItem(idx)

    params = Split(CStr(entry(2)), "|")
    ReDim args(LBound(params) To UBound(params))
    For i = LBound(params) To UBound(params)
        one = AskArgument(CStr(entry(0)), params(i), i + 1, UBound(params) + 1)
        If Len(one) = 0 Then Exit Sub
        ' An empty argument - "fx.RiskBetaλ(2, 3, , , ...)" - is how a LAMBDA
        ' sees an optional parameter as omitted.
        If modMonteCarlo.McParamIsOptional(params(i)) And modMonteCarlo.McIsNone(one) Then one = ""
        args(i) = one
    Next i

    If modMonteCarlo.McTrialsDefined(wb) Then
        trialsExpr = modMonteCarlo.MC_TRIALS_NAME
    Else
        trials = AskTrials()
        If trials < 1 Then Exit Sub
        trialsExpr = modMonteCarlo.McDefineTrials(wb, trials)
    End If

    Set target = AskTarget("Where should the trials go?" & vbCrLf & vbCrLf & _
                           "They spill DOWN from the cell you pick, so leave room below it.")
    If target Is Nothing Then Exit Sub
    If Not TargetIsClear(target) Then Exit Sub

    AppStateManager.FastModeOn

    If Not modMonteCarlo.McEnsureLibrary(wb, added) Then
        AppStateManager.FastModeOff
        Warn "The Monte Carlo functions could not be added to this workbook, so " & _
             "the formula was not written." & vbCrLf & vbCrLf & _
             "Try 'Install or Update Monte Carlo Library' on its own to see why."
        Exit Sub
    End If

    vid = modMonteCarlo.McNextVarId(wb)
    formulaText = modMonteCarlo.McBuildCall(CStr(entry(1)), args, trialsExpr, vid)
    modMonteCarlo.McWriteFormula target, formulaText

    AppStateManager.FastModeOff

    MsgBox CStr(entry(0)) & " written to " & target.Cells(1, 1).Address(False, False) & _
           " as VarID " & vid & "." & vbCrLf & vbCrLf & formulaText & vbCrLf & vbCrLf & _
           IIf(added > 0, "The library was installed into this workbook (" & added & _
                          " functions), so the file will keep working for people " & _
                          "who do not have XL Edge." & vbCrLf & vbCrLf, "") & _
           "Trial count comes from " & modMonteCarlo.MC_TRIALS_NAME & _
           ", which you can change in Name Manager.", _
           vbInformation, DIALOG_TITLE
    Exit Sub

Oops:
    AppStateManager.FastModeOff
    ShowError
End Sub


' --- shared worker for the two result helpers -------------------------------

Private Sub InsertFromSelection(ByVal baseName As String, ByVal hint As String)
    Dim wb As Workbook
    Dim source As Range, target As Range
    Dim ref As String, formulaText As String
    Dim added As Long

    On Error GoTo Oops
    If Not HaveWorkbook(wb) Then Exit Sub

    If Not modXLEdgeHelpers.GetSelectionRange(source, True, DIALOG_TITLE) Then Exit Sub
    If source Is Nothing Then
        Warn hint
        Exit Sub
    End If

    Set target = AskTarget("Where should the result go?")
    If target Is Nothing Then Exit Sub
    If Not TargetIsClear(target) Then Exit Sub

    ref = modMonteCarlo.McSpillReference(source, target)
    If Len(ref) = 0 Then
        Warn hint
        Exit Sub
    End If

    AppStateManager.FastModeOn
    If Not modMonteCarlo.McEnsureLibrary(wb, added) Then
        AppStateManager.FastModeOff
        Warn "The Monte Carlo functions could not be added to this workbook."
        Exit Sub
    End If
    formulaText = "=" & modMonteCarlo.McFunctionName(baseName) & "(" & ref & ")"
    modMonteCarlo.McWriteFormula target, formulaText
    AppStateManager.FastModeOff

    MsgBox "Written to " & target.Cells(1, 1).Address(False, False) & ":" & vbCrLf & vbCrLf & _
           formulaText, vbInformation, DIALOG_TITLE
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
' is not obvious from the error. Better to ask before writing.
Private Function TargetIsClear(ByVal target As Range) As Boolean
    Dim cell As Range
    Set cell = target.Cells(1, 1)

    If Len(cell.formula) = 0 Then
        TargetIsClear = True
        Exit Function
    End If
    TargetIsClear = modXLEdgeHelpers.ConfirmDestructive( _
        cell.Address(False, False) & " already contains something." & vbCrLf & vbCrLf & _
        "Overwrite it?", DIALOG_TITLE)
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
