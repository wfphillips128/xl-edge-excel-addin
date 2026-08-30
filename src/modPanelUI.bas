Attribute VB_Name = "modPanelUI"
Option Explicit

' ============================================================================
'  modPanelUI
'  The Tools menu item, what it reads off the user's selection, and the build
'  it runs when the dialog says go.
'
'  Everything user-facing lives here so the four modules below it stay pure:
'  modPanelSpec and modPanelFormulas touch no Excel at all, modPanelBuilder
'  writes cells, modPanelChart draws. This one is the only place that knows
'  there is a user.
'
'  WHY THE HOUSE OPENER IS NOT USED
'  --------------------------------
'  The standard first line for a range macro is
'
'      If Not GetSelectionRange(rg, True, TITLE_TXT) Then Exit Sub
'
'  and it is wrong here. It MsgBoxes and aborts when nothing is selected, and
'  in this tool "nothing selected" is a legitimate path - it means "give me a
'  placeholder block to paste over". So the selection is PROBED rather than
'  demanded, and GetSelectionRange is used only inside the branch that has
'  already decided there is a block worth reading.
' ============================================================================

Private Const TITLE_TXT As String = "Create Panel Chart"

' Chart size in points. The density limits in modPanelSpec are derived from
' these, less an allowance for the title, legend and axis labels, so raising
' the chart raises the limits with it.
'
' DEFERRED: the house rule is that a tunable belongs in tblConstants on
' shtReference, exposed as a Property Get in AddInStorage with a DEF_ fallback.
' These are left here for the first release deliberately - moving them means
' re-importing add-in-storage.bas and modSettingsUI.bas, which are shipped
' v2.01 code, and this feature is worth proving before it touches them.
Private Const PANEL_CHART_W As Double = 720
Private Const PANEL_CHART_H As Double = 500
Private Const PANEL_PLOT_W As Double = 720
Private Const PANEL_PLOT_H As Double = 400

' Below this share, a row or column is not a header.
Private Const HDR_SHARE As Double = 0.5


' What the selection turned out to be. Filled by DetectPanelSource, which
' never raises and never messages - it only reports.
Public Type TPanelSource
    hasBlock     As Boolean
    body         As Range         ' the numeric rectangle, headers stripped
    namesCol     As Range         ' panel names, or Nothing
    periodRow    As Range         ' period labels, or Nothing
    elementRow   As Range         ' element group names, or Nothing
    nPanels      As Long
    width        As Long
    periodsGuess As Long
    elemGuess    As Long
    textCells    As Long          ' body cells that are text, reported once
    note         As String        ' the sentence the form shows
    confident    As Boolean       ' False = the form must be looked at
End Type


' ============================================================================
'  Ribbon
' ============================================================================

' The Tools menu item. MUST be Public: a Private callback gives "callback not
' found" at load and Excel does not say which one.
Public Sub CreatePanelChart(control As IRibbonControl)
    LaunchPanelChart
End Sub


' ============================================================================
'  Launcher
' ============================================================================

' Also the Immediate-window entry point:  modPanelUI.LaunchPanelChart
Public Sub LaunchPanelChart()

    Dim f As frmPanelChart
    Dim sp As TPanelSpec, st As TPanelStyle
    Dim src As TPanelSource
    Dim errNum As Long, errTxt As String
    Dim problem As String

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open a workbook first - the panel chart is written to a new " & _
               "sheet in it.", vbExclamation, TITLE_TXT
        Exit Sub
    End If

    ' Tell the geometry how big the plot will be, so the density limits on the
    ' form are the ones this build will actually hit.
    PanelSetPlotExtent PANEL_PLOT_W, PANEL_PLOT_H

    DetectPanelSource src

    On Error GoTo Failed
    Set f = New frmPanelChart
    f.Preload src.note, src.hasBlock, src.periodsGuess, src.elemGuess, _
              src.nPanels, src.confident, _
              UniqueSheetName(ActiveWorkbook)
    ' NO FastModeOn around a modal dialog. House rule, and doubly justified
    ' here: the form reads no worksheet at all - every number on it comes from
    ' modPanelSpec, which is pure arithmetic. If FastModeOn were on and the
    ' form errored, EnableEvents would stay False for the rest of the Excel
    ' session across every open workbook, because an add-in changes GLOBAL
    ' state.
    f.Show vbModal

    If Not f.OK Then GoTo Done
    ReadFormInto f, sp, st, src

    On Error GoTo CleanUp
    AppStateManager.FastModeOn
    problem = BuildPanelSheet(ActiveWorkbook, sp, st, src, f.useSelection)

CleanUp:
    ' Read the error into locals BEFORE FastModeOff - reading Err after an
    ' intervening procedure call reads whatever that call left behind.
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, TITLE_TXT

    If errNum = 0 And Len(problem) > 0 Then
        MsgBox "The grid was built but did not verify:" & vbCrLf & vbCrLf & _
               problem, vbExclamation, TITLE_TXT
    End If

Done:
    On Error Resume Next
    Unload f
    Set f = Nothing
    Exit Sub

Failed:
    MsgBox "Could not open the panel chart dialog." & vbCrLf & vbCrLf & _
           Err.description, vbExclamation, TITLE_TXT
End Sub


' Marshal the form's scalars into the two Types.
'
' The form cannot hand back a TPanelSpec: a Public Type declared in a STANDARD
' module may not be a parameter or return type of a Public procedure in a form
' or class module - VBA refuses to compile it, with a message that does not
' explain itself. So everything crosses the boundary as a scalar and is
' reassembled here, on the standard-module side, where Types are legal.
Private Sub ReadFormInto(ByVal f As frmPanelChart, ByRef sp As TPanelSpec, _
                         ByRef st As TPanelStyle, ByRef src As TPanelSource)

    st = PanelStyleDefaults()

    sp.kind = f.KindIndex
    sp.rows = f.GridRows
    sp.cols = f.GridCols
    sp.Periods = f.Periods
    sp.Elements = f.Elements
    sp.SheetName = f.SheetName
    sp.title = f.TitleText
    sp.BandFrac = f.BandFrac
    sp.NTicks = f.NTicks
    sp.lead = f.LeadSlots
    sp.FloorHeadroom = f.FloorHeadroom

    Select Case f.IncludeZeroMode
        Case 0: sp.zeroSet = True: sp.includeZero = False
        Case 1: sp.zeroSet = True: sp.includeZero = True
        Case Else: sp.zeroSet = False          ' -1: whatever the kind wants
    End Select

    st.chartTitle = f.WantChartTitle
    st.valueTicks = f.WantValueTicks
    st.valueLabels = f.WantValueLabels

    ' Labels typed on the form win; otherwise PanelSpecInit fills defaults in.
    NamesFromCsv f.ElementNamesCsv, sp.elementNames, sp.Elements
    NamesFromCsv f.PeriodLabelsCsv, sp.periodLabels, sp.Periods

    ' Panel names come off the sheet when there is one to read.
    If f.useSelection And Not src.namesCol Is Nothing Then
        NamesFromRange src.namesCol, sp.panelNames, sp.rows * sp.cols
    End If
    If f.useSelection And Not src.periodRow Is Nothing And _
       Len(f.PeriodLabelsCsv) = 0 Then
        NamesFromRange src.periodRow, sp.periodLabels, sp.Periods
    End If
End Sub


' ============================================================================
'  The build
' ============================================================================

Private Function BuildPanelSheet(ByVal wb As Workbook, ByRef sp As TPanelSpec, _
                                 ByRef st As TPanelStyle, ByRef src As TPanelSource, _
                                 ByVal useSelection As Boolean) As String

    Dim ws As Worksheet, obj As ChartObject
    Dim ly As TPanelLayout
    Dim anns() As TAnnotation, anCount As Long
    Dim vals As Variant
    Dim problem As String

    If useSelection And src.hasBlock Then
        vals = ValuesForGrid(src, sp)
    Else
        vals = Empty                     ' the builder writes placeholder data
    End If

    ' Always a NEW sheet, deduped - which is why this tool needs no
    ' ConfirmDestructive prompt: there is nothing to overwrite.
    Set ws = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
    On Error Resume Next
    ws.name = sp.SheetName
    If Err.Number <> 0 Then
        Err.Clear
        ws.name = UniqueSheetName(wb)
    End If
    On Error GoTo 0
    sp.SheetName = ws.name

    problem = PanelWriteSheet(ws, sp, st, vals, ly, anns, anCount)
    If Len(problem) > 0 Then
        BuildPanelSheet = problem
        Exit Function
    End If

    Set obj = PanelDrawChart(ws, sp, st, ly, anns, anCount, vals, _
                             PANEL_CHART_W, PANEL_CHART_H)
    PanelPageSetup ws, obj

    problem = PanelVerify(ws, sp, ly, obj.Chart)

    ' Gridlines are a window property, not a sheet one, so the sheet has to be
    ' the active one to turn them off. This is the single place the tool
    ' activates anything, and the user wants to land on the result anyway.
    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    Err.Clear
    On Error GoTo 0

    BuildPanelSheet = problem
End Function


' Reshape the selected block into the engine's input rectangle.
'
' The grid may hold MORE cells than there are panels - a prime panel count
' leaves no factor pair, and the engine already draws nothing where a panel is
' absent. Spare rows are left blank rather than filled with zeros, because a
' blank draws nothing and a zero draws a flat panel that is not there.
Private Function ValuesForGrid(ByRef src As TPanelSource, _
                               ByRef sp As TPanelSpec) As Variant
    Dim raw As Variant, out() As Variant
    Dim r As Long, c As Long, nCells As Long

    raw = src.body.Value2
    nCells = PS_NPanels(sp)

    ReDim out(1 To nCells, 1 To sp.Elements * sp.Periods)

    For r = 1 To nCells
        For c = 1 To sp.Elements * sp.Periods
            If r <= UBound(raw, 1) And c <= UBound(raw, 2) Then
                If IsNumeric(raw(r, c)) And Not IsEmpty(raw(r, c)) Then
                    out(r, c) = raw(r, c)
                Else
                    out(r, c) = Empty
                End If
            Else
                out(r, c) = Empty
            End If
        Next c
    Next r

    ValuesForGrid = out
End Function


' ============================================================================
'  Selection detection
'
'  Reads the shape of what the user selected. Never raises, never messages -
'  a failure here just means the placeholder path, which is a legitimate
'  outcome rather than an error.
' ============================================================================

Public Sub DetectPanelSource(ByRef src As TPanelSource)

    Dim rg As Range, ws As Worksheet
    Dim arr As Variant
    Dim nR As Long, nC As Long
    Dim hasPeriodRow As Boolean, hasElemRow As Boolean, hasNamesCol As Boolean
    Dim r0 As Long, c0 As Long

    src.hasBlock = False
    src.confident = False
    src.note = "Nothing usable is selected, so a placeholder block will be " & _
               "written for you to paste over."

    ' -- rejections, in order. Any one of them means placeholder mode.
    If TypeName(Application.Selection) <> "Range" Then Exit Sub
    Set rg = Application.Selection
    If rg.Areas.Count > 1 Then Exit Sub
    On Error Resume Next
    Set ws = rg.Worksheet
    Err.Clear
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' Someone clicked a row or column header.
    If rg.rows.Count = ws.rows.Count Then Exit Sub
    If rg.Columns.Count = ws.Columns.Count Then Exit Sub
    ' Nobody builds a panel chart from that, and reading it would stall.
    If rg.Cells.Count > 65536 Then Exit Sub

    ' -- single-cell promotion. The common real case: the user clicks inside
    '    their table and hits the menu item.
    If rg.Cells.Count = 1 Then
        Set rg = rg.CurrentRegion
        If rg.Cells.Count <= 1 Then Exit Sub
        src.confident = False
    End If

    arr = rg.Value2
    If Not IsArray(arr) Then Exit Sub
    nR = UBound(arr, 1)
    nC = UBound(arr, 2)
    If nR < 1 Or nC < 1 Then Exit Sub

    ' -- header stripping. Only cells INSIDE the selection are inspected;
    '    reaching outside what the user chose is how a tool becomes
    '    unpredictable.
    If nR >= 3 Then
        If RowIsSparseText(arr, 1, nC) And RowIsMostlyText(arr, 2, nC) And _
           RowIsMostlyNumeric(arr, 3, nC) Then
            hasElemRow = True
            hasPeriodRow = True
        End If
    End If
    If Not hasPeriodRow And nR >= 2 Then
        If RowIsMostlyText(arr, 1, nC) And RowIsMostlyNumeric(arr, 2, nC) Then
            hasPeriodRow = True
        End If
    End If

    r0 = 1
    If hasElemRow Then r0 = r0 + 1
    If hasPeriodRow Then r0 = r0 + 1
    If r0 > nR Then Exit Sub

    c0 = 1
    If ColIsMostlyText(arr, 1, r0, nR) Then
        hasNamesCol = True
        c0 = 2
    End If
    If c0 > nC Then Exit Sub

    ' -- the block itself
    '
    ' ws.Range, NOT rg.Range. rg.Cells(r, c) already returns an ABSOLUTE cell,
    ' but Range.Range() re-interprets the addresses it is given as offsets from
    ' its OWN top-left - so rg.Range(rg.Cells(1, 1), ...) on a selection at C3
    ' resolves to E5, silently shifting the whole block down and right by the
    ' selection's own origin. The size stays correct, which is what makes it
    ' hard to spot: the panel count and column count both look right.
    Set src.body = ws.Range(ws.Cells(rg.Row + r0 - 1, rg.Column + c0 - 1), _
                            ws.Cells(rg.Row + nR - 1, rg.Column + nC - 1))
    src.nPanels = nR - r0 + 1
    src.width = nC - c0 + 1
    If src.nPanels < 1 Or src.width < 1 Then Exit Sub

    If hasNamesCol Then _
        Set src.namesCol = ws.Range(ws.Cells(rg.Row + r0 - 1, rg.Column), _
                                    ws.Cells(rg.Row + nR - 1, rg.Column))
    If hasPeriodRow Then _
        Set src.periodRow = ws.Range(ws.Cells(rg.Row + r0 - 2, rg.Column + c0 - 1), _
                                     ws.Cells(rg.Row + r0 - 2, rg.Column + nC - 1))
    If hasElemRow Then _
        Set src.elementRow = ws.Range(ws.Cells(rg.Row, rg.Column + c0 - 1), _
                                      ws.Cells(rg.Row, rg.Column + nC - 1))

    src.textCells = CountTextCells(arr, r0, nR, c0, nC)

    ' -- headers just OUTSIDE the selection
    '
    ' People select the DATA and leave the labels around it - which is exactly
    ' what the first real sheet did: a 15 x 12 block selected with the partner
    ' names one column left and Jan..Dec one row above, both outside it.
    ' Ignoring them gives a chart captioned "Panel 01" and "P1", which is no
    ' use to anyone.
    '
    ' So when a header was not found inside, look exactly ONE row up and ONE
    ' column left, accept it only if it reads as labels, and never look further
    ' than that. Reaching further is how a tool becomes unpredictable.
    Dim probe As Range

    ' Text names, or a column of distinct codes - cost centres, GL accounts
    ' and entity numbers are labels even though they are numeric.
    If src.namesCol Is Nothing And src.body.Column > 1 Then
        Set probe = src.body.offset(0, -1).Resize(src.body.rows.Count, 1)
        If RangeIsMostlyText(probe) Or RangeIsDistinctCodes(probe) Then
            Set src.namesCol = probe
        End If
    End If

    ' Text months, or a row of real dates. A row of dates sitting above a
    ' numeric block is a period header - nothing else it could be.
    If src.periodRow Is Nothing And src.body.Row > 1 Then
        Set probe = src.body.offset(-1, 0).Resize(1, src.body.Columns.Count)
        If RangeIsMostlyText(probe) Or RangeIsAllDates(probe) Then
            Set src.periodRow = probe
        End If
    End If

    ' An element-name row sits above the period row, its names spread across
    ' the groups with blanks between - that sparseness is what tells it apart
    ' from the period row, which is text all the way along.
    If src.elementRow Is Nothing And Not src.periodRow Is Nothing Then
        If src.periodRow.Row > 1 Then
            Set probe = src.periodRow.offset(-1, 0)
            If RangeIsSparseText(probe) Then Set src.elementRow = probe
        End If
    End If

    ' -- periods and elements, in order of how much they can be trusted
    InferPeriodsElements src

    src.hasBlock = True
    src.note = "Read from " & ws.name & "!" & src.body.Address & " - " & _
               src.nPanels & " panel row" & Plural(src.nPanels) & " x " & _
               src.width & " value column" & Plural(src.width) & " -> " & _
               src.elemGuess & " element" & Plural(src.elemGuess) & " x " & _
               src.periodsGuess & " period" & Plural(src.periodsGuess) & "."
    If src.textCells > 0 Then
        src.note = src.note & "  " & src.textCells & " cell" & _
                   Plural(src.textCells) & " in the block are text and will " & _
                   "be treated as empty."
    End If
End Sub


' Rows x cols is NOT inferred - the user is asked. 20 panels is equally 5x4,
' 4x5, 10x2 or 2x10, and only they know which one reads right.
'
' Periods and elements ARE inferred, in descending order of how far the
' evidence can be trusted. It works off the header RANGES, so it does not care
' whether they were found inside the selection or alongside it.
Private Sub InferPeriodsElements(ByRef src As TPanelSource)

    Dim e As Long, p As Long

    ' 1. From the element row. Exact, not a guess: the group names ARE the
    '    elements.
    If Not src.elementRow Is Nothing Then
        e = CountNonBlank(src.elementRow)
        If e >= 1 Then
            If src.width Mod e = 0 Then
                src.elemGuess = e
                src.periodsGuess = src.width \ e
                src.confident = True
                Exit Sub
            End If
        End If
    End If

    If Not src.periodRow Is Nothing Then

        ' 2. From REPEATING period labels. Jan..Dec Jan..Dec gives 12 at once.
        For p = 2 To src.width \ 2
            If src.width Mod p = 0 Then
                If LabelsMatch(src.periodRow, 1, 1 + p) Then
                    src.periodsGuess = p
                    src.elemGuess = src.width \ p
                    src.confident = True
                    Exit Sub
                End If
            End If
        Next p

        ' 3. All labels DISTINCT means one element group. There is nothing to
        '    be unsure about here - Jan..Dec with no repeat is twelve periods
        '    of one series - so do not tell the user to go and check it.
        If LabelsAllDistinct(src.periodRow) Then
            src.elemGuess = 1
            src.periodsGuess = src.width
            src.confident = True
            Exit Sub
        End If
    End If

    ' 4. No headers at all. One element is the only sensible default, but it IS
    '    a guess - twelve columns could be two elements of six - so say so.
    src.elemGuess = 1
    src.periodsGuess = src.width
    src.confident = False
End Sub


' ============================================================================
'  Self-check - the cheap gate, and the install smoke test
'
'      Immediate window:   ?modPanelUI.RunPanelSelfCheck
'
'  It needs no workbook and no chart. If this fails or is "not defined",
'  modPanelSpec did not import and nothing else is worth trying.
' ============================================================================

Public Function RunPanelSelfCheck() As String
    PanelSetPlotExtent PANEL_PLOT_W, PANEL_PLOT_H
    RunPanelSelfCheck = PanelSelfCheck()
End Function


' ============================================================================
'  Small helpers
' ============================================================================

' ----------------------------------------------------------------------------
'  Header tests on a Range
'
'  The arr-based tests above examine cells INSIDE the selection; these examine
'  a row or column beside it, where there is no array to index into.
' ----------------------------------------------------------------------------

' What a header cell READS AS, not what it stores.
'
' A date cell holds 46046 and displays "Jan-24"; the reader typed the latter
' and that is what belongs on the axis. This is the same trap the engine
' guards on the way out by setting NumberFormat "@" before writing - guarding
' one end and not the other would just move the 46046 rather than remove it.
'
' .Text is the displayed string, so it honours whatever format the user chose
' - "Jan-24", "Q1 24", "2024-01" all survive. Its one failure is a column too
' narrow to show the value, where it returns "#####"; that falls back to an
' explicit rendering rather than putting hashes on the chart.
Private Function CellLabel(ByVal c As Range) As String
    Dim t As String
    On Error Resume Next
    t = c.text
    Err.Clear
    On Error GoTo 0

    If Len(Trim$(t)) > 0 And InStr(1, t, "#") = 0 Then
        CellLabel = Trim$(t)
        Exit Function
    End If

    Dim v As Variant
    v = c.Value2
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsDate(c.value) Then
        CellLabel = Format$(c.value, "mmm-yy")
    Else
        CellLabel = Trim$(CStr(v))
    End If
End Function

' Every non-blank cell is a real date. Excel stores dates as numbers, so this
' asks the CELL, not the value - 46046 alone is just a number.
Private Function RangeIsAllDates(ByVal rg As Range) As Boolean
    Dim c As Range, seen As Long
    If rg Is Nothing Then Exit Function
    For Each c In rg.Cells
        If Not IsEmpty(c.Value2) Then
            If Not IsDate(c.value) Then Exit Function
            seen = seen + 1
        End If
    Next c
    RangeIsAllDates = (seen >= 2)
End Function

' A column of identifiers: fully populated and every value different.
'
' Deliberately strict. This column sits just outside a selection the user made
' on purpose, so it is probably a label - but it could be data they chose to
' exclude, and wrong panel titles read worse than obviously generic ones. One
' repeat, or one gap, and it is treated as data.
Private Function RangeIsDistinctCodes(ByVal rg As Range) As Boolean
    Dim c As Range, n As Long, i As Long, j As Long
    Dim vals() As String

    If rg Is Nothing Then Exit Function
    n = rg.Cells.Count
    If n < 2 Then Exit Function

    ReDim vals(1 To n)
    i = 0
    For Each c In rg.Cells
        i = i + 1
        vals(i) = CellLabel(c)
        If Len(vals(i)) = 0 Then Exit Function        ' a gap means data
    Next c

    For i = 1 To n - 1
        For j = i + 1 To n
            If StrComp(vals(i), vals(j), vbTextCompare) = 0 Then Exit Function
        Next j
    Next i

    RangeIsDistinctCodes = True
End Function

Private Function CountNonBlank(ByVal rg As Range) As Long
    Dim c As Range, n As Long
    If rg Is Nothing Then Exit Function
    For Each c In rg.Cells
        If IsTextCell(c.Value2) Then n = n + 1
    Next c
    CountNonBlank = n
End Function

' Mostly text AND not mostly numbers. A row of years is text-ish to a naive
' test but it is data, not a header.
Private Function RangeIsMostlyText(ByVal rg As Range) As Boolean
    Dim c As Range, txt As Long, num As Long, tot As Long
    If rg Is Nothing Then Exit Function
    For Each c In rg.Cells
        tot = tot + 1
        If IsTextCell(c.Value2) Then
            txt = txt + 1
        ElseIf IsNumCell(c.Value2) Then
            num = num + 1
        End If
    Next c
    If tot = 0 Then Exit Function
    RangeIsMostlyText = (txt >= tot * HDR_SHARE) And (num < tot * HDR_SHARE)
End Function

' Some text, with gaps - the shape of merged element-group names.
Private Function RangeIsSparseText(ByVal rg As Range) As Boolean
    Dim c As Range, txt As Long, tot As Long
    If rg Is Nothing Then Exit Function
    For Each c In rg.Cells
        tot = tot + 1
        If IsTextCell(c.Value2) Then txt = txt + 1
    Next c
    If tot = 0 Then Exit Function
    RangeIsSparseText = (txt >= 1) And (txt < tot * HDR_SHARE)
End Function

' Compare two 1-based positions along a single-row range.
Private Function LabelsMatch(ByVal rg As Range, ByVal i As Long, _
                             ByVal j As Long) As Boolean
    If rg Is Nothing Then Exit Function
    If i < 1 Or j < 1 Then Exit Function
    If i > rg.Cells.Count Or j > rg.Cells.Count Then Exit Function
    LabelsMatch = (StrComp(CellLabel(rg.Cells(1, i)), _
                           CellLabel(rg.Cells(1, j)), vbTextCompare) = 0)
End Function

' True when every non-blank label differs from every other. Repeats are what
' reveal several element groups; all-distinct means one.
Private Function LabelsAllDistinct(ByVal rg As Range) As Boolean
    Dim i As Long, j As Long, a As String, b As String, n As Long
    If rg Is Nothing Then Exit Function
    n = rg.Cells.Count
    For i = 1 To n - 1
        a = CellLabel(rg.Cells(1, i))
        If Len(a) > 0 Then
            For j = i + 1 To n
                b = CellLabel(rg.Cells(1, j))
                If Len(b) > 0 Then
                    If StrComp(a, b, vbTextCompare) = 0 Then Exit Function
                End If
            Next j
        End If
    Next i
    LabelsAllDistinct = True
End Function


Private Function Nz(ByVal v As Variant) As Variant
    If IsError(v) Then Nz = "" Else If IsEmpty(v) Then Nz = "" Else Nz = v
End Function

Private Function Plural(ByVal n As Long) As String
    If n = 1 Then Plural = "" Else Plural = "s"
End Function

Private Function IsTextCell(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then IsTextCell = (Len(Trim$(v)) > 0)
End Function

Private Function IsNumCell(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    IsNumCell = IsNumeric(v)
End Function

Private Function RowIsMostlyText(ByRef arr As Variant, ByVal r As Long, _
                                 ByVal nC As Long) As Boolean
    Dim c As Long, hits As Long
    For c = 1 To nC
        If IsTextCell(arr(r, c)) Then hits = hits + 1
    Next c
    RowIsMostlyText = (hits >= nC * HDR_SHARE)
End Function

Private Function RowIsMostlyNumeric(ByRef arr As Variant, ByVal r As Long, _
                                    ByVal nC As Long) As Boolean
    Dim c As Long, hits As Long
    For c = 1 To nC
        If IsNumCell(arr(r, c)) Then hits = hits + 1
    Next c
    RowIsMostlyNumeric = (hits >= nC * HDR_SHARE)
End Function

' An element-name row is SPARSE text: a name over each group, blanks between,
' because the group headers are merged. That sparseness is what tells it apart
' from the period-label row directly beneath, which is text all the way across.
Private Function RowIsSparseText(ByRef arr As Variant, ByVal r As Long, _
                                 ByVal nC As Long) As Boolean
    Dim c As Long, hits As Long
    For c = 1 To nC
        If IsTextCell(arr(r, c)) Then hits = hits + 1
    Next c
    RowIsSparseText = (hits >= 1 And hits < nC * HDR_SHARE)
End Function

Private Function ColIsMostlyText(ByRef arr As Variant, ByVal c As Long, _
                                 ByVal r0 As Long, ByVal nR As Long) As Boolean
    Dim r As Long, hits As Long, n As Long
    n = nR - r0 + 1
    If n < 1 Then Exit Function
    For r = r0 To nR
        If IsTextCell(arr(r, c)) Then hits = hits + 1
    Next r
    ColIsMostlyText = (hits >= n * HDR_SHARE)
End Function

' Text in the body is reported once, never silently coerced to zero.
Private Function CountTextCells(ByRef arr As Variant, ByVal r0 As Long, _
                                ByVal nR As Long, ByVal c0 As Long, _
                                ByVal nC As Long) As Long
    Dim r As Long, c As Long, n As Long
    For r = r0 To nR
        For c = c0 To nC
            If IsTextCell(arr(r, c)) Then n = n + 1
        Next c
    Next r
    CountTextCells = n
End Function


' Split a comma-separated list typed on the form into a name array. Left
' untouched when the count does not match, so PanelSpecInit's own defaults and
' its count check both still apply.
Private Sub NamesFromCsv(ByVal csv As String, ByRef out() As String, _
                         ByVal want As Long)
    Dim parts() As String, i As Long
    If Len(Trim$(csv)) = 0 Then Exit Sub
    parts = Split(csv, ",")
    If UBound(parts) - LBound(parts) + 1 <> want Then Exit Sub

    ReDim out(0 To want - 1)
    For i = 0 To want - 1
        out(i) = Trim$(parts(LBound(parts) + i))
    Next i
End Sub


' Read names off the sheet. Short ranges are padded with blanks rather than
' refused: a grid may legitimately hold more cells than panels, and a blank
' name means that panel simply prints no title.
Private Sub NamesFromRange(ByVal rg As Range, ByRef out() As String, _
                           ByVal want As Long)
    Dim c As Range, i As Long
    If rg Is Nothing Or want < 1 Then Exit Sub

    ReDim out(0 To want - 1)
    For i = 0 To want - 1
        out(i) = ""
    Next i

    ' CellLabel, not Value2: a date column would otherwise hand the chart
    ' 46046 instead of Jan-24. Short ranges leave the tail blank rather than
    ' failing - a grid may hold more cells than panels, and a blank name just
    ' means that panel prints no title.
    i = 0
    For Each c In rg.Cells
        If i > want - 1 Then Exit For
        out(i) = CellLabel(c)
        i = i + 1
    Next c
End Sub
