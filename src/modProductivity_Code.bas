Attribute VB_Name = "modProductivity_Code"
Option Explicit

' LOWER_BOUND, UPPER_BOUND, DEFAULT_ROW_HEIGHT, DEFAULT_COLUMN_WIDTH and
' myFooter moved to tblConstants on the add-in's reference sheet. They are
' still visible here as global Property Get procedures in AddInStorage, so
' every reference below is unchanged. Edit them from the XL Edge settings
' form, not in code.
'
' Note: myFooter now returns the "&<size>" point-size code plus the prose,
' assembled from two settings (myFooter and FOOTER_FONT_SIZE). The table
' stores readable text only.



Private Sub ApplyFactorToSelection(factor As Double)
'Multiplies the numeric CONSTANTS in the selection by factor. Formulas are left
'alone by construction -- SpecialCells(xlCellTypeConstants) never returns them.

    Const TITLE_TXT As String = "Scale Selection"

    Dim rng As Range
    Dim data As Variant
    Dim r As Long, c As Long
    Dim errNum As Long
    Dim errTxt As String

    If TypeName(Application.Selection) <> "Range" Then Exit Sub

    ' Restrict to numeric constants only
    On Error Resume Next
    Set rng = Application.Selection.SpecialCells(xlCellTypeConstants, xlNumbers)
    Err.Clear
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "No numeric constant cells found in selection.", vbExclamation, TITLE_TXT
        Exit Sub
    End If

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    data = rng.Value

    ' Handle single-cell case (not an array)
    If Not IsArray(data) Then
        If IsNumeric(data) Then data = data * factor
    Else
        For r = 1 To UBound(data, 1)
            For c = 1 To UBound(data, 2)
                If Not IsError(data(r, c)) Then
                    If IsNumeric(data(r, c)) And Not IsEmpty(data(r, c)) Then
                        data(r, c) = data(r, c) * factor
                    End If
                End If
            Next c
        Next r
    End If

    rng.Value = data

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, TITLE_TXT
End Sub


Sub MultiplyBy1000(control As IRibbonControl)
    ApplyFactorToSelection 1000
End Sub



Sub DivideBy1000(control As IRibbonControl)
    ApplyFactorToSelection 0.001
End Sub


Sub MultiplyByUserInput(control As IRibbonControl)

    Dim factor As Variant

    factor = Application.InputBox("Enter multiplication factor:", Type:=1)

    ' Handle Cancel
    If factor = False Then Exit Sub

    ' Block zero or blank-equivalent
    If factor = 0 Then
        MsgBox "Multiplication factor must be non-zero.", vbExclamation
        Exit Sub
    End If

    ApplyFactorToSelection 1 * factor

End Sub


Sub DivideByUserInput(control As IRibbonControl)

    Dim divisor As Variant

    divisor = Application.InputBox("Enter divisor:", Type:=1)

    ' Handle Cancel
    If divisor = False Then Exit Sub

    ' Block zero
    If divisor = 0 Then
        MsgBox "Divisor must be non-zero.", vbExclamation
        Exit Sub
    End If

    ApplyFactorToSelection 1 / divisor

End Sub


Public Sub FormatSelectedColumns(control As IRibbonControl)
    Dim errNum As Long
    Dim errTxt As String

    Dim ws As Worksheet
    Dim rngSel As Range
    Dim targetRange As Range
    Dim rngUsed As Range
    
    Dim formatChoice As Variant
    Dim modeChoice As Variant
    Dim formatString As String
    

    ' Validate selection
    If Not TypeOf Application.Selection Is Range Then Exit Sub
    Set rngSel = Application.Selection
    Set ws = rngSel.Worksheet

    ' -------------------------
    ' MODE TOGGLE (default = RANGE)
    ' -------------------------
    modeChoice = Application.InputBox( _
        prompt:="Select mode:" & vbCrLf & _
                "1 = Column Mode (used rows only)" & vbCrLf & _
                "2 = Range Mode (default)", _
        Title:="Formatting Mode", _
        Type:=1)

    ' Default to Range Mode if Cancel or blank
    If modeChoice = False Or modeChoice = 0 Then
        modeChoice = 2
    End If

    If modeChoice <> 1 And modeChoice <> 2 Then
        MsgBox "Invalid selection. Choose 1 or 2.", vbExclamation
        Exit Sub
    End If

    ' -------------------------
    ' FORMAT SELECTION
    ' -------------------------
    formatChoice = Application.InputBox( _
        prompt:="Select format:" & vbCrLf & _
                "1 = Comma" & vbCrLf & _
                "2 = Currency" & vbCrLf & _
                "3 = Percent" & vbCrLf & _
                "4 = Multiple" & vbCrLf & _
                "5 = Variance" & vbCrLf & _
                "6 = Variance Percent" & vbCrLf & _
                "7 = Factor" & vbCrLf & _
                "8 = Date" & vbCrLf & _
                "9 = Year", _
        Title:="Format Selection", _
        Type:=1)
        
    If formatChoice = False Then Exit Sub

    Select Case CLng(formatChoice)
        Case 1
            formatString = "_(#,##0_);(#,##0);_(""-""_);_(@_)"
        Case 2
            formatString = "_($#,##0_);($#,##0);_(""-""_);_(@_)"
        Case 3
            formatString = "_(#,##0.0%_);(#,##0.0%);_(#,##0.0%_);_(@_)"
        Case 4
            formatString = "_(0x_);(0x);_(""-""_);_(@_)"
        Case 5
            formatString = "[Color10]" & ChrW(&H25B2) & "_(#,##0_);[Red]" & ChrW(&H25BC) & "(#,##0);_(""-""_);_(@_)"
        Case 6
            formatString = "[Color10]" & ChrW(&H25B2) & "_(#,##0.0%_);[Red]" & ChrW(&H25BC) & "(#,##0.0%);_(#,##0.0%_);_(@_)"
        Case 7
            formatString = "_(###0.0000_);(###0.0000);_(""-""_);_(@_)"
        Case 8
            formatString = "mm/dd/yyyy;@"
        Case 9
            formatString = "0000"
        Case Else
            MsgBox "Invalid format selection.", vbExclamation
            Exit Sub
    End Select

    ' -------------------------
    ' BUILD TARGET RANGE
    ' -------------------------
    If modeChoice = 1 Then
        ' ? Column Mode (USED ROWS ONLY — optimized)
        Set rngUsed = ws.UsedRange
        Set targetRange = Intersect(rngSel.EntireColumn, rngUsed)
        
        If targetRange Is Nothing Then Exit Sub
    Else
        ' ? Default Range Mode
        Set targetRange = rngSel
    End If

    ' -------------------------
    ' APPLY FORMAT
    ' -------------------------

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    targetRange.NumberFormat = formatString

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, "Format Selected Column"

End Sub



Sub ColumnWidthsPaste(control As IRibbonControl)

    Dim src As Range
    Dim destTopLeft As Range
    Dim dest As Range
    Dim i As Long
    Dim srcWS As Worksheet, destWS As Worksheet
    
    On Error GoTo CleanFail
    
    '-----------------------------
    ' 1) Validate source selection
    '-----------------------------
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell range first.", vbExclamation, "Copy Column Widths"
        Exit Sub
    End If
    
    Set src = Selection
    
    ' Disallow multi-area selections (e.g., Ctrl+Click disjoint ranges)
    If src.Areas.Count > 1 Then
        MsgBox "Please select a single, contiguous range (not multiple areas).", _
               vbExclamation, "Copy Column Widths"
        Exit Sub
    End If
    
    ' Guard against empty selection (rare, but defensive)
    If src.Cells.CountLarge = 0 Then
        MsgBox "Your selection appears to be empty.", vbExclamation, "Copy Column Widths"
        Exit Sub
    End If
    
    Set srcWS = src.Worksheet
    
    '---------------------------------------------
    ' 2) Prompt for destination top-left cell only
    '---------------------------------------------
    ' Application.InputBox with Type:=8 returns a Range object.
    Set destTopLeft = Application.InputBox( _
        prompt:="Select the TOP-LEFT cell of the destination range." & vbCrLf & _
                "The destination range will automatically resize to match the source selection (" & _
                src.rows.Count & " rows x " & src.Columns.Count & " columns).", _
        Title:="Copy Column Widths (No Clipboard)", _
        Type:=8)
    
    ' If user clicks Cancel, InputBox returns False (as a Variant),
    ' but because we're assigning to a Range, it raises an error.
    ' We handle cancel via error handler below and a specific check here.
    If destTopLeft Is Nothing Then
        Exit Sub
    End If
    
    Set destWS = destTopLeft.Worksheet
    
    '---------------------------------------------------
    ' 3) Build destination range to match source dimensions
    '---------------------------------------------------
    ' Ensure destination will not exceed worksheet bounds.
    If destTopLeft.Row + src.rows.Count - 1 > destWS.rows.Count Then
        MsgBox "Destination range would extend past the bottom of the worksheet.", _
               vbCritical, "Copy Column Widths"
        Exit Sub
    End If
    
    If destTopLeft.Column + src.Columns.Count - 1 > destWS.Columns.Count Then
        MsgBox "Destination range would extend past the right edge of the worksheet.", _
               vbCritical, "Copy Column Widths"
        Exit Sub
    End If
    
    Set dest = destWS.Range(destTopLeft, _
                            destWS.Cells(destTopLeft.Row + src.rows.Count - 1, _
                                         destTopLeft.Column + src.Columns.Count - 1))
    
    '----------------------------------------
    ' 4) Prevent overlap (only meaningful on same sheet)
    '----------------------------------------
    If srcWS Is destWS Then
        If Not Intersect(src, dest) Is Nothing Then
            MsgBox "Source and destination ranges overlap. Please choose a different destination.", _
                   vbCritical, "Copy Column Widths"
            Exit Sub
        End If
    End If
    
    '----------------------------------------
    ' 5) Copy column widths WITHOUT clipboard
    '----------------------------------------
    AppStateManager.FastModeOn
    
    ' Copy each column width from source to destination.
    ' Note: ColumnWidth is a property of the entire column (not just the cells),
    ' and Excel applies it at the column level on the destination worksheet.
    For i = 1 To src.Columns.Count
        destWS.Columns(destTopLeft.Column + i - 1).ColumnWidth = _
            srcWS.Columns(src.Column + i - 1).ColumnWidth
    Next i
    
    AppStateManager.FastModeOff
    
    MsgBox "Column widths copied successfully (no clipboard used).", _
           vbInformation, "Copy Column Widths"
    Exit Sub

CleanFail:
    ' Handle common cancel case and unexpected errors defensively.
    AppStateManager.FastModeOff
    
    If Err.Number = 424 Or Err.Number = 13 Or Err.Number = 1004 Then
        ' These are common when InputBox is cancelled or invalid selection is made.
        ' We'll exit quietly or with a gentle notice depending on context.
        ' If cancelled, Err 424/13/1004 may occur depending on Excel version/context.
        Exit Sub
    End If
    
    MsgBox "An unexpected error occurred:" & vbCrLf & _
           Err.Number & " - " & Err.description, _
           vbCritical, "Copy Column Widths"

End Sub


Sub ExactFormulaCopy(control As IRibbonControl)

    Dim src As Range
    Dim destTopLeft As Range
    Dim dest As Range
    Dim srcWS As Worksheet, destWS As Worksheet
    Dim r As Long, c As Long
    Dim srcCell As Range, destCell As Range
    Dim response As VbMsgBoxResult

    On Error GoTo CleanFail

    '-----------------------------
    ' 1) Validate source selection
    '-----------------------------
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell range first.", vbExclamation, "Copy Formulas Exactly"
        Exit Sub
    End If

    Set src = Selection

    ' Require a single contiguous block (no multi-area selections)
    If src.Areas.Count > 1 Then
        MsgBox "Please select a single, contiguous range (not multiple separate areas).", _
               vbExclamation, "Copy Formulas Exactly"
        Exit Sub
    End If

    If src.Cells.CountLarge = 0 Then
        MsgBox "Your selection appears to be empty.", vbExclamation, "Copy Formulas Exactly"
        Exit Sub
    End If

    ' Optional safety guard: discourage gigantic selections (defensive)
    ' Adjust threshold as desired.
    If src.Cells.CountLarge > 200000 Then
        response = MsgBox("The selected range is very large (" & Format$(src.Cells.CountLarge, "#,##0") & " cells)." & vbCrLf & _
                          "This may take a while. Continue?", _
                          vbQuestion + vbYesNo, "Copy Formulas Exactly")
        If response <> vbYes Then Exit Sub
    End If

    Set srcWS = src.Worksheet

    '-------------------------------------------------
    ' 2) Prompt for destination TOP-LEFT cell (Type:=8)
    '-------------------------------------------------
    Set destTopLeft = Application.InputBox( _
        prompt:="Select the TOP-LEFT cell of the destination." & vbCrLf & _
                "Destination will be resized to match source (" & src.rows.Count & " rows x " & src.Columns.Count & " columns).", _
        Title:="Copy Formulas Exactly (No Clipboard)", _
        Type:=8)

    ' If user cancels, Application.InputBox(Type:=8) often raises an error,
    ' which is handled in CleanFail. This check is extra defensive.
    If destTopLeft Is Nothing Then Exit Sub

    Set destWS = destTopLeft.Worksheet

    '-------------------------------------------------------
    ' 3) Build destination range to match source dimensions
    '    and ensure it stays within worksheet boundaries
    '-------------------------------------------------------
    If destTopLeft.Row + src.rows.Count - 1 > destWS.rows.Count Then
        MsgBox "Destination range would extend past the bottom of the worksheet.", _
               vbCritical, "Copy Formulas Exactly"
        Exit Sub
    End If

    If destTopLeft.Column + src.Columns.Count - 1 > destWS.Columns.Count Then
        MsgBox "Destination range would extend past the right edge of the worksheet.", _
               vbCritical, "Copy Formulas Exactly"
        Exit Sub
    End If

    Set dest = destWS.Range(destTopLeft, _
                            destWS.Cells(destTopLeft.Row + src.rows.Count - 1, _
                                         destTopLeft.Column + src.Columns.Count - 1))

    '----------------------------------------
    ' 4) Prevent overlap (only on same sheet)
    '----------------------------------------
    If srcWS Is destWS Then
        If Not Intersect(src, dest) Is Nothing Then
            MsgBox "Source and destination ranges overlap." & vbCrLf & _
                   "Please choose a different destination.", _
                   vbCritical, "Copy Formulas Exactly"
            Exit Sub
        End If
    End If

    '---------------------------------------------------------
    ' 5) If destination has anything in it, confirm overwrite
    '---------------------------------------------------------
    ' CountA counts non-empty cells including cells with formulas (even if they display "").
    If Application.WorksheetFunction.CountA(dest) > 0 Then
        response = MsgBox("The destination range contains existing data (values and/or formulas)." & vbCrLf & _
                          "Do you want to overwrite it?", _
                          vbQuestion + vbYesNo, "Confirm Overwrite")
        If response <> vbYes Then Exit Sub
    End If

    '-----------------------------------------------
    ' 6 & 7) Copy formulas literally, else clear cell
    '-----------------------------------------------
    AppStateManager.FastModeOn

    ' Note:
    ' - We assign .Formula cell-by-cell to preserve the literal A1-style text.
    ' - We do NOT use .Copy or clipboard.
    ' - If a source cell has no formula, destination cell is cleared.
    For r = 1 To src.rows.Count
        For c = 1 To src.Columns.Count

            Set srcCell = src.Cells(r, c)
            Set destCell = dest.Cells(r, c)

            If srcCell.HasFormula Then
                ' Copy formula string exactly as returned by .Formula.
                ' This keeps the literal text (e.g. "=A1" stays "=A1").
                destCell.formula = srcCell.formula
            Else
                ' Source is not a formula -> clear destination per requirement.
                destCell.ClearContents
            End If

        Next c
    Next r

    AppStateManager.FastModeOff

    MsgBox "Formulas copied successfully (no clipboard used)." & vbCrLf & _
           "Non-formula source cells cleared in destination.", _
           vbInformation, "Copy Formulas Exactly"

    Exit Sub

CleanFail:
    ' Ensure UI settings are restored even if something goes wrong.
    AppStateManager.FastModeOff

    ' Common cancel/selection errors from InputBox(Type:=8) or invalid range selection.
    Select Case Err.Number
        Case 424, 13, 1004
            ' Exit quietly on cancel or common selection-related errors.
            Exit Sub
        Case Else
            MsgBox "An unexpected error occurred:" & vbCrLf & _
                   Err.Number & " - " & Err.description, _
                   vbCritical, "Copy Formulas Exactly"
    End Select

End Sub


Public Sub ChangeSumToSubTotal(control As IRibbonControl)
    ' Convert leading =SUM(...) formulas in the selection to =SUBTOTAL(9,...).
    ' Selection scope. Only touches cells whose formula begins exactly with "=SUM(".
    On Error GoTo CleanExit

    If TypeName(Selection) <> "Range" Then Exit Sub

    ' Confirm BEFORE touching app state, while the screen is still live.
    Dim answer As String
    answer = Application.InputBox( _
        "Confirm that you would like to change" & vbLf & _
        "all highlighted formulas from SUM to SUBTOTAL (Y/N).", _
        "Enter Y/N:", "Y")
    If UCase$(Trim$(answer)) <> "Y" Then Exit Sub

    ' Narrow to just the formula cells in the selection: skips blanks/constants and
    ' makes whole-column selections cheap. SpecialCells raises 1004 if none exist.
    Dim formulaCells As Range
    On Error Resume Next
    Set formulaCells = Selection.SpecialCells(xlCellTypeFormulas)
    On Error GoTo CleanExit
    If formulaCells Is Nothing Then Exit Sub        ' no formulas in selection

    ' --- capture prior app state, then enter fast mode ---
    Dim errNum As Long
    Dim errTxt As String

    AppStateManager.FastModeOn

    Dim area As Range, f As Variant
    Dim r As Long, c As Long
    For Each area In formulaCells.Areas             ' handles multi-area selections
        f = area.formula                             ' ONE read per area
        If Not IsArray(f) Then                        ' single-cell area -> scalar
            If UCase$(Left$(CStr(f), 5)) = "=SUM(" Then _
                area.formula = "=SUBTOTAL(9," & Mid$(CStr(f), 6)
        Else
            For r = 1 To UBound(f, 1)
                For c = 1 To UBound(f, 2)
                    If UCase$(Left$(CStr(f(r, c)), 5)) = "=SUM(" Then _
                        f(r, c) = "=SUBTOTAL(9," & Mid$(CStr(f(r, c)), 6)
                Next c
            Next r
            area.formula = f                          ' ONE write-back per area
        End If
    Next area

CleanExit:
    ' FastModeOff is a no-op when FastModeOn never ran, so the early Exit Sub
    ' paths above cost nothing here.
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, "Change SUM to SUBTOTAL"
End Sub



Sub InsertCompanyName(control As IRibbonControl)
    ' Cycles the active cell through the company names in tblCompany.
    '
    ' The list used to be a hardcoded Array() here. It now comes from the
    ' add-in's reference sheet so a user can edit it from the settings form.
    '
    ' AddInStorage.CompanyList() returns a 1-BASED array with blanks already
    ' stripped -- unlike the old Array(), which was 0-based with a deliberate
    ' empty first element. Hence no "skip the blank on wrap" special case here.

    If TypeName(Selection) <> "Range" Then Exit Sub

    Dim options As Variant
    options = AddInStorage.CompanyList()
    If Not AddInStorage.HasItems(options) Then
        MsgBox "No company names are set up yet." & vbCrLf & vbCrLf & _
               "Add them under XL Edge settings.", vbInformation, "Insert Company Name"
        Exit Sub
    End If

    Dim current As String
    current = CStr(ActiveCell.Value)

    ' Find the current value and advance one, wrapping at the end.
    Dim i As Long
    For i = LBound(options) To UBound(options)
        If StrComp(current, CStr(options(i)), vbTextCompare) = 0 Then
            If i = UBound(options) Then
                ActiveCell.Value = options(LBound(options))
            Else
                ActiveCell.Value = options(i + 1)
            End If
            Exit Sub
        End If
    Next i

    ' Cell is blank or holds something not in the list -> start at the top.
    ActiveCell.Value = options(LBound(options))

End Sub


Sub FormatLightUnderlinetoSelectedRange(control As IRibbonControl)
'Format all cells in the selected range

Dim intArea As Long
Dim rngCell As Range
 
    For intArea = 1 To Selection.Areas.Count
        For Each rngCell In Selection.Areas(intArea).Cells
            With rngCell.Borders(xlEdgeBottom)
                .LineStyle = xlDash
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        Next
    Next

End Sub


Sub UnmergeSelectedSheets(control As IRibbonControl)
'Unmerges every merged cell on the selected sheets.

    Const TITLE_TXT As String = "Unmerge Selected Sheets"

    Dim sh As Object
    Dim startCell As Range
    Dim errNum As Long
    Dim errTxt As String

    If ActiveWindow Is Nothing Then Exit Sub
    If TypeName(Application.Selection) = "Range" Then Set startCell = ActiveCell

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    For Each sh In ActiveWindow.SelectedSheets
        If TypeOf sh Is Worksheet Then sh.Cells.UnMerge
    Next sh

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    If Not startCell Is Nothing Then Application.GoTo startCell
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, TITLE_TXT
End Sub


Sub UnmergeAllSheets(control As IRibbonControl)
'Unmerges every merged cell on every worksheet in the active workbook.

    Const TITLE_TXT As String = "Unmerge All Sheets"

    Dim ws As Worksheet
    Dim errNum As Long
    Dim errTxt As String

    If ActiveWorkbook Is Nothing Then Exit Sub

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.UnMerge
    Next ws

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, TITLE_TXT
End Sub


Sub UnmergedCellsToCenterAcross(control As IRibbonControl)
'Converts horizontal merged cells on the active sheet to Center Across
'Selection, which looks the same and behaves far better.

    Const TITLE_TXT As String = "Unmerge and Center Across"

    Dim ws As Worksheet
    Dim rng As Range
    Dim c As Range
    Dim mergedArea As Range
    Dim processed As Object      ' Dictionary to avoid reprocessing same merged area
    Dim addr As String
    Dim errNum As Long
    Dim errTxt As String

    ' Ensure we're on a worksheet
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub

    Set ws = ActiveSheet
    Set rng = ws.UsedRange
    Set processed = CreateObject("Scripting.Dictionary")

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    For Each c In rng
        If c.MergeCells Then
            Set mergedArea = c.MergeArea

            ' Only handle horizontal merges (one row high)
            If mergedArea.rows.Count = 1 Then
                addr = mergedArea.Address

                ' Process each merged area only once
                If Not processed.Exists(addr) Then
                    processed.Add addr, True

                    mergedArea.UnMerge
                    mergedArea.HorizontalAlignment = xlCenterAcrossSelection
                End If
            End If
        End If
    Next c

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, TITLE_TXT
End Sub


Sub Wrap_with_Round(control As IRibbonControl)
'Wraps every selected formula in ROUND(..., n).
    Const TITLE_TXT As String = "Wrap with Round"

    Dim v As Variant

    ' Type:=1 makes Excel police the input as a number for us. Cancel returns
    ' Boolean False -- the old version concatenated that straight into the
    ' formula and produced =ROUND(...,False).
    v = Application.InputBox( _
            prompt:="Enter the number of decimal places" & vbCrLf & _
                    "to use in the ROUND formula:", _
            Title:=TITLE_TXT, Default:="0", Type:=1)

    If VarType(v) = vbBoolean Then Exit Sub

    WrapSelectionFormulas TITLE_TXT, "=ROUND(", "," & CStr(CLng(v)) & ")", "=ROUND"
End Sub


Sub Wrap_with_Error(control As IRibbonControl)
'Wraps every selected formula in IFERROR(..., value).
    Const TITLE_TXT As String = "Wrap with IFERROR"

    Dim v As Variant

    v = Application.InputBox( _
            prompt:="Enter the value to return" & vbCrLf & _
                    "if an error condition is found:" & vbCrLf & vbCrLf & _
                    "(text needs its own quotes, e.g. ""n/a"" )", _
            Title:=TITLE_TXT, Default:="0")

    If VarType(v) = vbBoolean Then Exit Sub
    If Len(Trim$(CStr(v))) = 0 Then Exit Sub

    WrapSelectionFormulas TITLE_TXT, "=IFERROR(", "," & CStr(v) & ")", "=IFERROR"
End Sub


Sub Wrap_Parenthesis(control As IRibbonControl)
'Wraps every selected formula in a pair of parentheses.
    Const TITLE_TXT As String = "Wrap with Parenthesis"

    If Not ConfirmProceed("Wrap all selected formulas with parentheses?", _
                          TITLE_TXT) Then Exit Sub

    WrapSelectionFormulas TITLE_TXT, "=(", ")", "=("
End Sub

Sub Wrap_Flip_Sign(control As IRibbonControl)
'Wraps every selected formula in a negation.
    Const TITLE_TXT As String = "Wrap with Flip Sign"

    If Not ConfirmProceed("Flip the sign on all selected formulas?", _
                          TITLE_TXT) Then Exit Sub

    WrapSelectionFormulas TITLE_TXT, "=-(", ")", "=-("
End Sub


Sub ChangeSignforInputCells(control As IRibbonControl)
'Flips the sign of the selected non-formula numeric cells.
    Const TITLE_TXT As String = "Change Sign of Inputs"

    If Not ConfirmProceed("Flip the sign on the selected non-formula cells?", _
                          TITLE_TXT) Then Exit Sub

    ApplySignToInputCells TITLE_TXT, False
End Sub


Sub RemoveNegativeSigns(control As IRibbonControl)
'Makes the selected non-formula numeric cells positive.
    Const TITLE_TXT As String = "Remove Negative Signs"

    If Not ConfirmProceed("Remove negative signs from the selected " & _
                          "non-formula cells?", TITLE_TXT) Then Exit Sub

    ApplySignToInputCells TITLE_TXT, True
End Sub


Sub InsertSymbol(control As IRibbonControl)
    Dim tempItem As String
    Dim cel As Range
    Dim celLen As Long
    Dim symbColorIndex As Long
    Dim symbFontSize As Long
    Dim selectedRange As Range
    Dim choice As Variant
    Dim symbolName As String
    
    ' Prompt user to choose symbol
    choice = Application.InputBox( _
        prompt:="Choose a symbol to insert (1 to 6):" & vbCrLf & _
                "  1 - Checkmark" & vbCrLf & _
                "  2 - Cross" & vbCrLf & _
                "  3 - Lambda" & vbCrLf & _
                "  4 - Warning" & vbCrLf & _
                "  5 - Delta" & vbCrLf & _
                "  6 - Sigma (Total)", _
        Title:="Insert Symbol", _
        Type:=1)    ' Type 1 = number only

    ' Handle Cancel or invalid input
    If choice = False Then Exit Sub          ' User pressed Cancel
    If choice < 1 Or choice > 6 Then
        MsgBox "Please enter a number from 1 to 6.", vbExclamation, "Invalid Choice"
        Exit Sub
    End If
    
    ' Map numeric choice to the existing symbol names
    Select Case choice
        Case 1: symbolName = "Checkmark"
        Case 2: symbolName = "Cross"
        Case 3: symbolName = "Lambda"
        Case 4: symbolName = "Warning"
        Case 5: symbolName = "Delta"
        Case 6: symbolName = "Sigma"
    End Select
    
    Set selectedRange = Application.Selection
    
    Select Case symbolName
        Case "Checkmark"
            symbColorIndex = 10
            symbFontSize = 14
            tempItem = ChrW(10004)
        Case "Cross"
            symbColorIndex = 3
            symbFontSize = 14
            tempItem = ChrW(10006)
        Case "Lambda"
            symbColorIndex = 10
            symbFontSize = 14
            tempItem = ChrW(955)
        Case "Warning"
            symbColorIndex = 3
            symbFontSize = 14
            tempItem = ChrW(9888)
        Case "Delta"
            symbColorIndex = 10
            symbFontSize = 14
            tempItem = ChrW(8710)
        Case "Sigma"
            symbColorIndex = 10
            symbFontSize = 14
            tempItem = ChrW(8721)
    End Select
    
    ' --------- LOOP PRESERVES EXISTING FORMATTING ---------
    Dim startPos As Long
    
    For Each cel In selectedRange.Cells
        With cel
            ' Length of current text (unformatted text only)
            celLen = Len(.text)   ' or .Value2; .Text matches what's displayed
            
            ' Position where the new symbol should be inserted
            startPos = celLen + 1
            
            ' Append the symbol as a new character
            .Characters(Start:=startPos, Length:=1).text = tempItem
            
            ' Format ONLY the last character (the symbol)
            .Characters(Start:=startPos, Length:=1).Font.ColorIndex = symbColorIndex
            .Characters(Start:=startPos, Length:=1).Font.Size = symbFontSize
        End With
    Next cel
End Sub


Sub Change_Formula_Absolute(control As IRibbonControl)
'Locks every reference in the selected formulas (the F4 effect).
'
'ConvertFormula's 4th argument is left as True, exactly as it was. True is not
'one of the XlReferenceType constants, but it is what this macro has always
'passed and what it has always been tested against -- changing it to
'xlAbsolute is a separate decision, not part of a state-handling refactor.

    Const TITLE_TXT As String = "Formulas to Absolute Refs"

    Dim rg As Range
    Dim cell As Range
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(rg, False, TITLE_TXT) Then Exit Sub

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    For Each cell In rg.Cells
        If cell.HasFormula Then
            ' A formula ConvertFormula cannot handle skips that cell rather than
            ' abandoning the run. That is what the old blanket On Error Resume
            ' Next around the whole procedure was really for -- but it also hid
            ' every other failure, including the ones worth seeing.
            On Error Resume Next
            cell.formula = Application.ConvertFormula(cell.formula, xlA1, xlA1, True)
            Err.Clear
            On Error GoTo CleanUp
        End If
    Next cell

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, TITLE_TXT
End Sub

Sub Change_Formula_Relative(control As IRibbonControl)
'Removes every locked reference from the selected formulas.

    Const TITLE_TXT As String = "Formulas to Relative Refs"

    Dim rg As Range
    Dim cell As Range
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(rg, False, TITLE_TXT) Then Exit Sub

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    For Each cell In rg.Cells
        If cell.HasFormula Then
            On Error Resume Next
            cell.formula = Application.ConvertFormula(cell.formula, xlA1, xlA1, xlRelative)
            Err.Clear
            On Error GoTo CleanUp
        End If
    Next cell

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, TITLE_TXT
End Sub


Sub Switch_Landscape(control As IRibbonControl)
'Landscape orientation on every selected worksheet. See the page setup notes
'at the foot of this module for what changed here and why.
    ApplyPageOrientation xlLandscape, "Switch Landscape"
End Sub

Sub Switch_Portrait(control As IRibbonControl)
'Portrait orientation on every selected worksheet.
    ApplyPageOrientation xlPortrait, "Switch Portrait"
End Sub

Sub StandardPageSetup(control As IRibbonControl)
'The house page setup, applied to every selected worksheet.

    Const TITLE_TXT As String = "Std Page Setup"

    Dim sh As Object
    Dim errNum As Long
    Dim errTxt As String

    If ActiveWindow Is Nothing Then Exit Sub

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    For Each sh In ActiveWindow.SelectedSheets
        If TypeOf sh Is Worksheet Then
            With sh.PageSetup
                .LeftMargin = Application.InchesToPoints(0.25)
                .RightMargin = Application.InchesToPoints(0.25)
                .TopMargin = Application.InchesToPoints(0.75)
                .BottomMargin = Application.InchesToPoints(0.5)
                .HeaderMargin = Application.InchesToPoints(0.5)
                .FooterMargin = Application.InchesToPoints(0.25)
                .CenterHorizontally = True
                .CenterVertically = False
                .Orientation = xlPortrait
                .Zoom = False
                .PaperSize = xlPaperLetter
                .FitToPagesWide = 1
                .FitToPagesTall = 1
            End With
        End If
    Next sh

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, TITLE_TXT
End Sub


Sub ClearFooters(control As IRibbonControl)
'Blanks all three footer zones on every selected worksheet.

    Const TITLE_TXT As String = "Clear Footers"

    Dim sh As Object
    Dim errNum As Long
    Dim errTxt As String

    If ActiveWindow Is Nothing Then Exit Sub

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    For Each sh In ActiveWindow.SelectedSheets
        If TypeOf sh Is Worksheet Then
            With sh.PageSetup
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
            End With
        End If
    Next sh

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, TITLE_TXT
End Sub

Sub UpdateLeftFooter(control As IRibbonControl)
'Full file path on the left, date and time on the right.

    Const TITLE_TXT As String = "Update Left Footer"

    Dim sh As Object
    Dim footerPath As String
    Dim errNum As Long
    Dim errTxt As String

    If ActiveWindow Is Nothing Then Exit Sub
    If ActiveWorkbook Is Nothing Then Exit Sub

    ' Read once, outside the loop. FullName is empty for a workbook that has
    ' never been saved, in which case the name is more use than nothing.
    footerPath = ActiveWorkbook.FullName
    If Len(footerPath) = 0 Then footerPath = ActiveWorkbook.name

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    For Each sh In ActiveWindow.SelectedSheets
        If TypeOf sh Is Worksheet Then
            With sh.PageSetup
                .LeftFooter = footerPath
                .RightFooter = "&D, &T"
            End With
        End If
    Next sh

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, TITLE_TXT
End Sub

Sub UpdateRightFooter(control As IRibbonControl)
'Date and time on the left, full file path on the right.

    Const TITLE_TXT As String = "Update Right Footer"

    Dim sh As Object
    Dim footerPath As String
    Dim errNum As Long
    Dim errTxt As String

    If ActiveWindow Is Nothing Then Exit Sub
    If ActiveWorkbook Is Nothing Then Exit Sub

    footerPath = ActiveWorkbook.FullName
    If Len(footerPath) = 0 Then footerPath = ActiveWorkbook.name

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    For Each sh In ActiveWindow.SelectedSheets
        If TypeOf sh Is Worksheet Then
            With sh.PageSetup
                .LeftFooter = "&D, &T"
                .RightFooter = footerPath
            End With
        End If
    Next sh

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, TITLE_TXT
End Sub

Sub CopySelectedWorksheetsToNewWorkbook(control As IRibbonControl)
'Copy selected worksheets to a new file and then remove all formulas

ActiveWorkbook.Windows(1).SelectedSheets.Copy

    Dim ws As Worksheet
    For Each ws In Worksheets
      With ws.UsedRange
        .Value = .Value
      End With
    Next ws


End Sub


Sub RemoveEmptyRows(control As IRibbonControl)
'Deletes every entirely blank ROW inside the current selection.
'
'Scope is the selection, which is what the ribbon supertip has always claimed.
'The previous version worked on ActiveSheet.UsedRange regardless of what was
'selected, mutated its own For counter, and raised a type mismatch the moment
'it met a cell holding an error value.

    Const TITLE_TXT As String = "Remove Empty Rows"

    Dim rg As Range
    Dim clipped As Range
    Dim unionRows As Range
    Dim i As Long
    Dim nRows As Long
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(rg, True, TITLE_TXT) Then Exit Sub

    If rg.rows.Count > 20000 Then
        Set clipped = Intersect(rg, rg.Worksheet.UsedRange)
        If clipped Is Nothing Then
            MsgBox "The selection holds no data, so there is nothing to remove.", _
                   vbInformation, TITLE_TXT
            Exit Sub
        End If
        Set rg = clipped
    End If

    For i = 1 To rg.rows.Count
        If IsBlankRange(rg.rows(i)) Then
            If unionRows Is Nothing Then
                Set unionRows = rg.rows(i).EntireRow
            Else
                Set unionRows = Union(unionRows, rg.rows(i).EntireRow)
            End If
            nRows = nRows + 1
        End If
    Next i

    If nRows = 0 Then
        MsgBox "No completely blank rows were found in the selection.", _
               vbInformation, TITLE_TXT
        Exit Sub
    End If

    If Not ConfirmDestructive( _
        "Delete " & nRows & " entire row(s)?" & vbCrLf & vbCrLf & _
        "Rows are judged blank using the selected cells only, but the whole sheet " & _
        "row is removed -- anything outside the selection on them will also be lost." & _
        vbCrLf & vbCrLf & "This cannot be undone.", TITLE_TXT) Then Exit Sub

    On Error GoTo CleanExit
    AppStateManager.FastModeOn
    unionRows.Delete

CleanExit:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    If errNum <> 0 Then
        MsgBox "Could not finish removing blank rows." & vbCrLf & vbCrLf & _
               "Error " & errNum & ": " & errTxt, vbExclamation, TITLE_TXT
    End If
End Sub

Sub replaceBlankWithZero(control As IRibbonControl)
'Puts a zero into every blank cell of the selection.
'
'The previous version opened with Selection.Value = Selection.Value, which
'silently converted every formula in the selection to a value before it did
'anything else. Nothing on the button said so.

    Const TITLE_TXT As String = "Replace Blanks with Zeros"

    Dim rg As Range
    Dim blanks As Range
    Dim textCells As Range
    Dim c As Range
    Dim n As Long
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(rg, False, TITLE_TXT) Then Exit Sub

    On Error GoTo CleanExit
    AppStateManager.FastModeOn

    On Error Resume Next
    Set blanks = rg.SpecialCells(xlCellTypeBlanks)
    Err.Clear
    Set textCells = rg.SpecialCells(xlCellTypeConstants, xlTextValues)
    Err.Clear
    On Error GoTo CleanExit

    If Not blanks Is Nothing Then
        blanks.Value = 0
        n = n + blanks.Cells.Count
    End If

    ' The old macro also zeroed cells holding a single space, so keep that.
    If Not textCells Is Nothing Then
        For Each c In textCells.Cells
            If Len(Trim$(CStr(c.Value2))) = 0 Then
                c.Value = 0
                n = n + 1
            End If
        Next c
    End If

CleanExit:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    If errNum <> 0 Then
        MsgBox "Could not finish." & vbCrLf & vbCrLf & _
               "Error " & errNum & ": " & errTxt, vbExclamation, TITLE_TXT
    ElseIf n = 0 Then
        MsgBox "There are no blank cells in the selection.", vbInformation, TITLE_TXT
    End If
End Sub

Sub removeChar(control As IRibbonControl)
'Removes the given text from every cell in the selection.
'
'The previous version called Selection.Replace INSIDE a For Each over the
'selection, so the whole replace ran once per cell -- on a 10,000 cell
'selection that is 10,000 full passes.

    Const TITLE_TXT As String = "Remove Char(s)"

    Dim rg As Range
    Dim rc As String
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(rg, False, TITLE_TXT) Then Exit Sub

    rc = InputBox("Character(s) to remove from the selected cells:", TITLE_TXT)
    If Len(rc) = 0 Then Exit Sub          ' Cancel, or nothing typed

    On Error GoTo CleanExit
    AppStateManager.FastModeOn

    rg.Replace What:=rc, Replacement:="", _
               LookAt:=xlPart, MatchCase:=False

CleanExit:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    If errNum <> 0 Then
        MsgBox "Could not finish." & vbCrLf & vbCrLf & _
               "Error " & errNum & ": " & errTxt, vbExclamation, TITLE_TXT
    End If
End Sub


Sub RemoveSpaces(control As IRibbonControl)
'Trims the text cells in the selection.
'
'The button says "Trim Spaces", so it now trims the way a spreadsheet person
'means it: leading and trailing spaces removed AND runs of spaces inside the
'text collapsed to one. The old code called VBA's Trim(), which only does the
'ends -- "Acme    Corp" came back unchanged.
'
'The old confirmation was a case-sensitive Y/N input box: typing a lower case
'y silently did nothing at all.

    Const TITLE_TXT As String = "Trim Spaces"

    Dim rg As Range
    Dim textCells As Range
    Dim area As Range
    Dim arr As Variant
    Dim r As Long, c As Long
    Dim n As Long
    Dim before As String
    Dim after As String
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(rg, False, TITLE_TXT) Then Exit Sub

    If Not ConfirmProceed( _
        "Remove excess spaces from the text cells in the selection?" & vbCrLf & vbCrLf & _
        "Leading and trailing spaces are removed, and runs of spaces inside the " & _
        "text are reduced to a single space.", TITLE_TXT) Then Exit Sub

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    ' Text constants only: numbers, dates and formulas are left alone.
    On Error Resume Next
    Set textCells = rg.SpecialCells(xlCellTypeConstants, xlTextValues)
    Err.Clear
    On Error GoTo CleanUp

    If Not textCells Is Nothing Then
        For Each area In textCells.Areas
            arr = area.Value2

            If IsArray(arr) Then
                For r = 1 To UBound(arr, 1)
                    For c = 1 To UBound(arr, 2)
                        before = CStr(arr(r, c))
                        after = CleanSpaces(before)
                        If after <> before Then n = n + 1
                        arr(r, c) = after
                    Next c
                Next r
                area.Value2 = arr
            Else
                before = CStr(arr)
                after = CleanSpaces(before)
                If after <> before Then n = n + 1
                area.Value2 = after
            End If
        Next area
    End If

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    If errNum <> 0 Then
        ReportError errNum, errTxt, TITLE_TXT
    Else
        MsgBox Format$(n, "#,##0") & " cell(s) changed.", vbInformation, TITLE_TXT
    End If
End Sub


Sub ClearExcessRowsAndColumns(control As IRibbonControl)
'Deletes every entirely blank ROW and COLUMN inside the current selection.
'
'Blankness is judged using the selected cells only, but the delete removes the
'whole sheet row / column -- so anything sitting outside the selection on those
'rows and columns goes too. That is what the confirmation is warning about.
'
'The previous version of this macro did something else entirely: it walked
'every worksheet in the workbook clearing formatting BEYOND the last used cell.
'That job is already covered by Tools > Shrink Excel File (ExcelShrinkFile).

    Const TITLE_TXT As String = "Remove Empty Rows and Cols"

    Dim rg As Range
    Dim clipped As Range
    Dim unionRows As Range
    Dim unionCols As Range
    Dim i As Long
    Dim nRows As Long
    Dim nCols As Long
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(rg, True, TITLE_TXT) Then Exit Sub

    'A whole-column or whole-row selection would otherwise mean a million
    'CountA calls. Clip to the used range first -- everything outside it is
    'blank by definition, and deleting a million empty rows achieves nothing.
    If rg.rows.Count > 20000 Or rg.Columns.Count > 2000 Then
        Set clipped = Intersect(rg, rg.Worksheet.UsedRange)
        If clipped Is Nothing Then
            MsgBox "The selection holds no data, so there is nothing to remove.", _
                   vbInformation, TITLE_TXT
            Exit Sub
        End If
        Set rg = clipped
    End If

    'Pass one: collect only. Deleting as we go would shift every index that
    'has not been tested yet, which is what makes the naive version skip rows.
    For i = 1 To rg.rows.Count
        If IsBlankRange(rg.rows(i)) Then
            If unionRows Is Nothing Then
                Set unionRows = rg.rows(i).EntireRow
            Else
                Set unionRows = Union(unionRows, rg.rows(i).EntireRow)
            End If
            nRows = nRows + 1
        End If
    Next i

    For i = 1 To rg.Columns.Count
        If IsBlankRange(rg.Columns(i)) Then
            If unionCols Is Nothing Then
                Set unionCols = rg.Columns(i).EntireColumn
            Else
                Set unionCols = Union(unionCols, rg.Columns(i).EntireColumn)
            End If
            nCols = nCols + 1
        End If
    Next i

    If nRows = 0 And nCols = 0 Then
        MsgBox "No completely blank rows or columns were found in the selection.", _
               vbInformation, TITLE_TXT
        Exit Sub
    End If

    If Not ConfirmDestructive( _
        "Delete " & nRows & " entire row(s) and " & nCols & " entire column(s)?" & _
        vbCrLf & vbCrLf & _
        "Rows and columns are judged blank using the selected cells only, but the " & _
        "whole sheet row and column is removed -- anything outside the selection " & _
        "on them will also be lost." & vbCrLf & vbCrLf & _
        "This cannot be undone.", TITLE_TXT) Then Exit Sub

    On Error GoTo CleanExit
    AppStateManager.FastModeOn

    'One Delete call per set. Deleting a union in a single call is far faster
    'than looping, and avoids the shifting-index problem completely.
    If Not unionRows Is Nothing Then unionRows.Delete
    If Not unionCols Is Nothing Then unionCols.Delete

CleanExit:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    If errNum <> 0 Then
        MsgBox "Could not finish removing blank rows and columns." & vbCrLf & vbCrLf & _
               "Error " & errNum & ": " & errTxt, vbExclamation, TITLE_TXT
    End If
End Sub




Public Sub FillBlankCellsWithValue(control As IRibbonControl)
    Dim errNum As Long
    Dim errTxt As String

    Dim rngSel As Range
    Dim blankCells As Range
    Dim inputValue As Variant
    
    ' Validate selection
    If Not TypeOf Application.Selection Is Range Then
        MsgBox "No valid range is selected.", vbCritical, "Error"
        Exit Sub
    End If
    
    Set rngSel = Application.Selection

    ' Prompt user
    inputValue = Application.InputBox( _
        prompt:="Enter value to insert into blank cells:", _
        Title:="Fill Blanks", _
        Type:=2)

    ' Cancel returns Boolean False. Testing the value itself meant a typed
    ' zero compared equal to False and was treated as a cancel.
    If VarType(inputValue) = vbBoolean Then Exit Sub

    ' Try to get blanks
    On Error Resume Next
    Set blankCells = rngSel.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    If Not blankCells Is Nothing Then
        ' Case 1: Real blanks inside UsedRange
        blankCells.Value = inputValue
    Else
        ' Case 2: Entire selection is empty (outside UsedRange)
        rngSel.Value = inputValue
    End If

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, "Fill Blanks with Value"

    MsgBox "Blank cells updated.", vbInformation

End Sub

Public Sub FillSelectionWithUniformRandom(control As IRibbonControl)
    Dim errNum As Long
    Dim errTxt As String
    Dim rng As Range
    Dim area As Range
    Dim arr As Variant
    Dim r As Long, c As Long
    
    ' Validate selection
    If Not TypeOf Application.Selection Is Range Then Exit Sub
    Set rng = Application.Selection

    Randomize

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    ' Process all cells (formulas included by design)
    For Each area In rng.Areas
        arr = area.Value2

        If IsArray(arr) Then
            For r = 1 To UBound(arr, 1)
                For c = 1 To UBound(arr, 2)
                    arr(r, c) = GetUniformRandom(LOWER_BOUND, UPPER_BOUND)
                Next c
            Next r
            area.Value2 = arr
        Else
            area.Value2 = GetUniformRandom(LOWER_BOUND, UPPER_BOUND)
        End If
    Next area

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, "Fill Range with Random Uniform Values"
End Sub

Public Sub FillSelectionWithNormalRandom(control As IRibbonControl)
    Dim errNum As Long
    Dim errTxt As String
    Dim rng As Range
    Dim area As Range
    Dim arr As Variant
    Dim r As Long, c As Long
    
    ' Validate selection
    If Not TypeOf Application.Selection Is Range Then Exit Sub
    Set rng = Application.Selection

    Randomize

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    ' Process all cells (formulas included by design)
    For Each area In rng.Areas
        arr = area.Value2

        If IsArray(arr) Then
            For r = 1 To UBound(arr, 1)
                For c = 1 To UBound(arr, 2)
                    arr(r, c) = GetNormalRandom(LOWER_BOUND, UPPER_BOUND)
                Next c
            Next r
            area.Value2 = arr
        Else
            area.Value2 = GetNormalRandom(LOWER_BOUND, UPPER_BOUND)
        End If
    Next area

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, "Fill Range with Random Normal Values"
End Sub


Public Function GetUniformRandom(lowerB As Double, upperB As Double) As Double
    GetUniformRandom = lowerB + (upperB - lowerB) * Rnd
End Function



Public Function GetNormalRandom(lowerB As Double, upperB As Double) As Double
    Dim u1 As Double, u2 As Double
    Dim z As Double
    Dim mean As Double, stdDev As Double

    mean = (lowerB + upperB) / 2
    stdDev = (upperB - lowerB) / 6

    ' Avoid log(0)
    Do
        u1 = Rnd
    Loop While u1 = 0

    u2 = Rnd

    z = Sqr(-2 * Log(u1)) * Cos(2 * WorksheetFunction.Pi() * u2)

    GetNormalRandom = mean + z * stdDev

    ' Clamp to bounds
    If GetNormalRandom < lowerB Then GetNormalRandom = lowerB
    If GetNormalRandom > upperB Then GetNormalRandom = upperB
End Function


Public Sub AddPrefix(control As IRibbonControl)
    Dim errNum As Long
    Dim errTxt As String
    Dim rng As Range
    Dim textCells As Range
    Dim area As Range
    Dim arr As Variant
    Dim r As Long, c As Long
    Dim prefixValue As Variant
    
    ' Validate selection
    If Not TypeOf Application.Selection Is Range Then Exit Sub
    Set rng = Application.Selection

    ' Prompt user
    prefixValue = Application.InputBox( _
        prompt:="Enter prefix:", _
        Title:="Prefix", _
        Type:=2)
        
    ' Cancel returns Boolean False; a typed zero must not count as one.
    If VarType(prefixValue) = vbBoolean Then Exit Sub

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    On Error Resume Next
    ' ? TEXT VALUES ONLY (this excludes numbers entirely)
    Set textCells = rng.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0

    If Not textCells Is Nothing Then
        For Each area In textCells.Areas
            arr = area.Value2

            If IsArray(arr) Then
                For r = 1 To UBound(arr, 1)
                    For c = 1 To UBound(arr, 2)
                        If Len(arr(r, c)) > 0 Then
                            arr(r, c) = prefixValue & arr(r, c)
                        End If
                    Next c
                Next r
                area.Value2 = arr
            Else
                If Len(arr) > 0 Then
                    area.Value2 = prefixValue & arr
                End If
            End If
        Next area
    End If

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, "Add Prefix"
End Sub


Public Sub AddSuffix(control As IRibbonControl)
    Dim errNum As Long
    Dim errTxt As String
    Dim rng As Range
    Dim textCells As Range
    Dim area As Range
    Dim arr As Variant
    Dim r As Long, c As Long
    Dim suffixValue As Variant
    
    ' Validate selection
    If Not TypeOf Application.Selection Is Range Then Exit Sub
    Set rng = Application.Selection

    ' Prompt user
    suffixValue = Application.InputBox( _
        prompt:="Enter Suffix:", _
        Title:="Suffix", _
        Type:=2)
        
    ' Cancel returns Boolean False; a typed zero must not count as one.
    If VarType(suffixValue) = vbBoolean Then Exit Sub

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    On Error Resume Next
    ' ? TEXT CONSTANTS ONLY (skips numbers and formulas)
    Set textCells = rng.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0

    If Not textCells Is Nothing Then
        For Each area In textCells.Areas
            arr = area.Value2

            If IsArray(arr) Then
                For r = 1 To UBound(arr, 1)
                    For c = 1 To UBound(arr, 2)
                        If Len(arr(r, c)) > 0 Then
                            arr(r, c) = arr(r, c) & suffixValue
                        End If
                    Next c
                Next r
                area.Value2 = arr
            Else
                If Len(arr) > 0 Then
                    area.Value2 = arr & suffixValue
                End If
            End If
        Next area
    End If

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, "Add Suffix"
End Sub


Public Sub RemovePrefix(control As IRibbonControl)
    Dim errNum As Long
    Dim errTxt As String
    Dim rng As Range
    Dim textCells As Range
    Dim area As Range
    Dim arr As Variant
    Dim r As Long, c As Long
    Dim prefixValue As Variant
    Dim prefixLen As Long
    
    ' Validate selection
    If Not TypeOf Application.Selection Is Range Then Exit Sub
    Set rng = Application.Selection

    ' Prompt user
    prefixValue = Application.InputBox( _
        prompt:="Enter prefix to remove:", _
        Title:="Remove Prefix", _
        Type:=2)
        
    ' Cancel returns Boolean False; a typed zero must not count as one.
    If VarType(prefixValue) = vbBoolean Then Exit Sub
    prefixLen = Len(prefixValue)

    If prefixLen = 0 Then Exit Sub

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    On Error Resume Next
    ' ? Text constants only
    Set textCells = rng.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0

    If Not textCells Is Nothing Then
        For Each area In textCells.Areas
            arr = area.Value2

            If IsArray(arr) Then
                For r = 1 To UBound(arr, 1)
                    For c = 1 To UBound(arr, 2)
                        If Len(arr(r, c)) >= prefixLen Then
                            If Left$(arr(r, c), prefixLen) = prefixValue Then
                                arr(r, c) = Mid$(arr(r, c), prefixLen + 1)
                            End If
                        End If
                    Next c
                Next r
                area.Value2 = arr
            Else
                If Len(arr) >= prefixLen Then
                    If Left$(arr, prefixLen) = prefixValue Then
                        area.Value2 = Mid$(arr, prefixLen + 1)
                    End If
                End If
            End If
        Next area
    End If

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, "Remove Prefix"
End Sub

Public Sub RemoveSuffix(control As IRibbonControl)
    Dim errNum As Long
    Dim errTxt As String
    Dim rng As Range
    Dim textCells As Range
    Dim area As Range
    Dim arr As Variant
    Dim r As Long, c As Long
    Dim suffixValue As Variant
    Dim suffixLen As Long
    
    ' Validate selection
    If Not TypeOf Application.Selection Is Range Then Exit Sub
    Set rng = Application.Selection

    ' Prompt user
    suffixValue = Application.InputBox( _
        prompt:="Enter suffix to remove:", _
        Title:="Remove Suffix", _
        Type:=2)
        
    ' Cancel returns Boolean False; a typed zero must not count as one.
    If VarType(suffixValue) = vbBoolean Then Exit Sub
    suffixLen = Len(suffixValue)

    If suffixLen = 0 Then Exit Sub

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    On Error Resume Next
    ' ? Text constants only
    Set textCells = rng.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0

    If Not textCells Is Nothing Then
        For Each area In textCells.Areas
            arr = area.Value2

            If IsArray(arr) Then
                For r = 1 To UBound(arr, 1)
                    For c = 1 To UBound(arr, 2)
                        If Len(arr(r, c)) >= suffixLen Then
                            If Right$(arr(r, c), suffixLen) = suffixValue Then
                                arr(r, c) = Left$(arr(r, c), Len(arr(r, c)) - suffixLen)
                            End If
                        End If
                    Next c
                Next r
                area.Value2 = arr
            Else
                If Len(arr) >= suffixLen Then
                    If Right$(arr, suffixLen) = suffixValue Then
                        area.Value2 = Left$(arr, Len(arr) - suffixLen)
                    End If
                End If
            End If
        Next area
    End If

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, "Remove Suffix"
End Sub


Sub ConvertSentenceCase(control As IRibbonControl)
    ConvertSentenceCase2
End Sub


Sub ConvertSentenceCase2()
Dim rng As Range
    For Each rng In Selection
    If WorksheetFunction.IsText(rng) Then
    rng.Value = UCase(Left(rng, 1)) & LCase(Right(rng, Len(rng) - 1))
    End If
    Next rng
End Sub


Sub ConvertProperCase(control As IRibbonControl)
    ConvertProperCase2
End Sub


Sub ConvertProperCase2()
Dim rng As Range
    For Each rng In Selection
    If WorksheetFunction.IsText(rng) Then
    rng.Value = WorksheetFunction.Proper(rng.Value)
    End If
    Next
End Sub

Sub ConvertLowerCase(control As IRibbonControl)
    ConvertLowerCase2
End Sub

Sub ConvertLowerCase2()
Dim rng As Range
    For Each rng In Selection
    If Application.WorksheetFunction.IsText(rng) Then
    rng.Value = LCase(rng)
    End If
    Next
End Sub

Sub ConvertUpperCase(control As IRibbonControl)
    ConvertUpperCase2
End Sub

Sub ConvertUpperCase2()
Dim rng As Range
    For Each rng In Selection
    If Application.WorksheetFunction.IsText(rng) Then
    rng.Value = UCase(rng)
    End If
    Next
End Sub

Sub ToggleCase(control As IRibbonControl)

Dim cell As Range
Dim FirstCellValue As String
Dim FirstCellValue2 As String
Dim FirstCellValue3 As String
Dim TestValue As Long

Set cell = Selection
FirstCellValue = cell.Cells(1, 1).Value
FirstCellValue2 = Left(FirstCellValue, 1)
FirstCellValue3 = Right(FirstCellValue, Len(FirstCellValue) - 1)


If FirstCellValue = LCase(FirstCellValue) Then
    TestValue = 4
ElseIf FirstCellValue = UCase(FirstCellValue) Then
    TestValue = 3
ElseIf (FirstCellValue2 = UCase(FirstCellValue2) And FirstCellValue3 = LCase(FirstCellValue3)) Then
    TestValue = 1
Else
    TestValue = 2
End If

    With Selection
    Select Case TestValue
        Case 1
            ConvertProperCase2
        Case 2
            ConvertUpperCase2
        Case 3
            ConvertLowerCase2
        Case 4
            ConvertSentenceCase2
        Case Else
            ConvertSentenceCase2
    End Select
    End With

End Sub


Sub ShowStartupPath(control As IRibbonControl)

    Dim startupPath As String
    startupPath = Application.startupPath
    
    ' Show message
    MsgBox "Personal.xlsb is located at:" & vbCrLf & startupPath, vbInformation, "Startup Path"
    
    ' Copy to clipboard (no external dependency)
    On Error GoTo ClipboardFail
    
    With CreateObject("htmlfile")
        .parentWindow.clipboardData.setData "text", startupPath
    End With
    
    Exit Sub

ClipboardFail:
    MsgBox "Path copied failed, but location is shown above." & vbCrLf & _
           "Error: " & Err.description, vbExclamation

End Sub


Sub ShowAddInLocation(control As IRibbonControl)

    Dim addInsPath As String
    addInsPath = Application.UserLibraryPath
    
    ' Show message
    MsgBox "Excel Add-ins folder is located at:" & vbCrLf & addInsPath, _
           vbInformation, "Add-Ins Path"
    
    ' Copy to clipboard (htmlfile method)
    On Error GoTo ClipboardFail
    
    With CreateObject("htmlfile")
        .parentWindow.clipboardData.setData "text", addInsPath
    End With
    
    Exit Sub

ClipboardFail:
    MsgBox "Path copy failed, but location is shown above." & vbCrLf & _
           "Error: " & Err.description, vbExclamation

End Sub


Sub RemoveFormulasfromSelection(control As IRibbonControl)
'Replaces every formula in the selection with its current value.
    Const TITLE_TXT As String = "Remove Formulas from Selection"

    Dim rg As Range
    Dim cellsDone As Long
    Dim areasFailed As Long
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(rg, False, TITLE_TXT) Then Exit Sub

    If Not ConfirmDestructive( _
        "Replace every formula in the selected range with its current value?" & _
        vbCrLf & vbCrLf & "This cannot be undone.", TITLE_TXT) Then Exit Sub

    On Error GoTo CleanExit
    AppStateManager.FastModeOn
    ConvertFormulasToValues rg, cellsDone, areasFailed

CleanExit:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportFormulaRemoval TITLE_TXT, cellsDone, areasFailed, vbNullString, errNum, errTxt
End Sub


Sub RemoveFormulasfromWorksheet(control As IRibbonControl)
'Replaces every formula on the active worksheet with its current value.
    Const TITLE_TXT As String = "Remove Formulas from Worksheet"

    Dim ws As Worksheet
    Dim cellsDone As Long
    Dim areasFailed As Long
    Dim errNum As Long
    Dim errTxt As String

    If ActiveWorkbook Is Nothing Then Exit Sub

    If Not TypeOf ActiveSheet Is Worksheet Then
        MsgBox "The active sheet is not a worksheet.", vbExclamation, TITLE_TXT
        Exit Sub
    End If
    Set ws = ActiveSheet

    If ws.ProtectContents Then
        MsgBox "'" & ws.name & "' is protected, so its formulas cannot be replaced.", _
               vbExclamation, TITLE_TXT
        Exit Sub
    End If

    If Not ConfirmDestructive( _
        "Replace every formula on '" & ws.name & "' with its current value?" & _
        vbCrLf & vbCrLf & "This cannot be undone.", TITLE_TXT) Then Exit Sub

    On Error GoTo CleanExit
    AppStateManager.FastModeOn
    ConvertFormulasToValues ws.UsedRange, cellsDone, areasFailed

CleanExit:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportFormulaRemoval TITLE_TXT, cellsDone, areasFailed, vbNullString, errNum, errTxt
End Sub

Sub RemoveFormulasfromWorkbook(control As IRibbonControl)
'Replaces every formula in the active workbook with its current value.
    Const TITLE_TXT As String = "Remove Formulas from Workbook"

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cellsDone As Long
    Dim areasFailed As Long
    Dim skipped As String
    Dim errNum As Long
    Dim errTxt As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook

    If Not ConfirmDestructive( _
        "Replace every formula in '" & wb.name & "' with its current value?" & _
        vbCrLf & vbCrLf & _
        "Every worksheet in the workbook is affected. This cannot be undone.", _
        TITLE_TXT) Then Exit Sub

    On Error GoTo CleanExit
    AppStateManager.FastModeOn

    For Each ws In wb.Worksheets
        If ws.ProtectContents Then
            skipped = skipped & vbCrLf & "    " & ws.name
        Else
            ConvertFormulasToValues ws.UsedRange, cellsDone, areasFailed
        End If
    Next ws

CleanExit:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportFormulaRemoval TITLE_TXT, cellsDone, areasFailed, skipped, errNum, errTxt
End Sub


Private Sub ApplyRowHeight(ByVal target As Range, ByVal rowHeight As Double)
    If target Is Nothing Then Exit Sub
    target.rowHeight = rowHeight
End Sub

Private Sub ApplyColumnWidth(ByVal target As Range, ByVal colWidth As Double)
    If target Is Nothing Then Exit Sub
    target.ColumnWidth = colWidth
End Sub


Sub SetRowHeights(control As IRibbonControl)

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a valid range.", vbExclamation
        Exit Sub
    End If

    ApplyRowHeight Selection, DEFAULT_ROW_HEIGHT

End Sub

Sub SetColumnWidths(control As IRibbonControl)

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a valid range.", vbExclamation
        Exit Sub
    End If

    ApplyColumnWidth Selection, DEFAULT_COLUMN_WIDTH

End Sub


Sub SetRowHeightsWorksheet(control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.UsedRange Is Nothing Then Exit Sub

    ApplyRowHeight ws.UsedRange, DEFAULT_ROW_HEIGHT

End Sub


Sub SetColumnWidthsWorksheet(control As IRibbonControl)
    Dim ws As Worksheet

    Set ws = ActiveSheet

    If ws.UsedRange Is Nothing Then Exit Sub

    ApplyColumnWidth ws.UsedRange, DEFAULT_COLUMN_WIDTH

End Sub


Sub ListFormulaToLeft(control As IRibbonControl)
'Writes a live "<--- =formula" listing of the cell immediately to the left.
    Const TITLE_TXT As String = "List Formula To Left"

    Dim c As Range

    If TypeName(Application.Selection) <> "Range" Then
        MsgBox "Please select a cell first.", vbExclamation, TITLE_TXT
        Exit Sub
    End If

    Set c = ActiveCell

    If c.Column = 1 Then
        MsgBox "No cell exists to the left.", vbExclamation, TITLE_TXT
        Exit Sub
    End If

    On Error GoTo ErrHandler

    If Not WriteFormulaListing(c) Then
        MsgBox "The cell to the left holds no formula, so there is nothing to list.", _
               vbInformation, TITLE_TXT
    End If
    Exit Sub

ErrHandler:
    MsgBox "Unable to retrieve formula text from the cell to the left." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.description, vbExclamation, TITLE_TXT
End Sub

Sub ListFormulaToLeftSelection(control As IRibbonControl)
'Same listing, applied to every cell in the selection.
    Const TITLE_TXT As String = "List Formula To Left"

    Dim rg As Range
    Dim c As Range
    Dim n As Long
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(rg, False, TITLE_TXT) Then Exit Sub

    On Error GoTo CleanExit
    AppStateManager.FastModeOn

    For Each c In rg.Cells
        If c.Column > 1 Then
            If WriteFormulaListing(c) Then n = n + 1
        End If
    Next c

CleanExit:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    If errNum <> 0 Then
        MsgBox "An error occurred while processing the selection." & vbCrLf & vbCrLf & _
               "Error " & errNum & ": " & errTxt, vbExclamation, TITLE_TXT
    ElseIf n = 0 Then
        MsgBox "None of the cells immediately to the left hold a formula, " & _
               "so nothing was listed.", vbInformation, TITLE_TXT
    End If
End Sub




Sub CenterAcrossSelectionCode(control As IRibbonControl)

    Dim rng As Range

    ' Validate selection
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a valid range.", vbExclamation
        Exit Sub
    End If

    Set rng = Selection

    ' Apply alignment
    rng.HorizontalAlignment = xlCenterAcrossSelection

End Sub



Sub DeleteCustomStyles_Workbook(control As IRibbonControl)

    Dim wb As Workbook
    Dim xStyle As Style
    Dim stylesToDelete As Collection
    Dim styleName As Variant
    Dim deleteCount As Long

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Set stylesToDelete = New Collection

    ' Collect custom styles first (avoid modifying collection during loop)
    For Each xStyle In wb.Styles
        If Not xStyle.BuiltIn Then
            stylesToDelete.Add xStyle.name
        End If
    Next xStyle

    On Error Resume Next  ' Some styles may fail if in use

    ' Delete collected styles
    For Each styleName In stylesToDelete
        wb.Styles(styleName).Delete
        If Err.Number = 0 Then deleteCount = deleteCount + 1
        Err.Clear
    Next styleName

    On Error GoTo 0

    MsgBox deleteCount & " custom styles deleted.", vbInformation, "Cleanup Complete"

End Sub



Sub SpeakCellContents(control As IRibbonControl)

    Dim rng As Range

    ' Validate selection
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a valid range.", vbExclamation
        Exit Sub
    End If

    Set rng = Selection

    On Error GoTo ErrHandler

    ' Speak selected cells
    rng.Speak

    Exit Sub

ErrHandler:
    MsgBox "Text-to-speech is unavailable or failed: " & Err.description, vbExclamation

End Sub


Sub RemoveSameSheetReferences(control As IRibbonControl)
'Strips "SheetName!" prefixes from formulas on the active sheet, where the
'reference points at the sheet the formula already lives on.

    Const TITLE_TXT As String = "Remove Same Sheet Refs"

    Dim sht As Worksheet
    Dim fndList As Variant
    Dim x As Long
    Dim errNum As Long
    Dim errTxt As String

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set sht = ActiveSheet

    fndList = Array("'" & sht.name & "'!", sht.name & "!")

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    For x = LBound(fndList) To UBound(fndList)
        sht.Cells.Replace What:=fndList(x), Replacement:="", _
            LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
            SearchFormat:=False, ReplaceFormat:=False
    Next x

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, TITLE_TXT
End Sub


Sub ExcelShrinkFile(control As IRibbonControl)
    Dim errNum As Long
    Dim errTxt As String

    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim fCell As Range
    Dim shp As Shape
    Dim shpLastRow As Long, shpLastCol As Long
    
    If ActiveWorkbook Is Nothing Then Exit Sub

    ' This used to walk ThisWorkbook -- the ADD-IN itself, not the file on
    ' screen. The button says "deletes unnecessary elements in the current
    ' Excel workbook", so it now does that. Because it deletes rows and
    ' columns in the user's own file, it asks first.
    If Not ConfirmDestructive( _
        "Delete all unused rows and columns from every sheet in '" & _
        ActiveWorkbook.name & "'?" & vbCrLf & vbCrLf & _
        "This resets each sheet's used range so the file saves smaller. " & _
        "It cannot be undone.", "Shrink Excel File") Then Exit Sub

    On Error GoTo CleanUp
    AppStateManager.FastModeOn
    ' FastModeOff restores whatever DisplayAlerts was before.
    Application.DisplayAlerts = False

    For Each ws In ActiveWorkbook.Worksheets
        With ws
            
            lastRow = 0
            lastCol = 0
            
            '--- Find last used cell (formulas + values)
            Set fCell = .Cells.Find(What:="*", LookIn:=xlFormulas, _
                                   SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            If Not fCell Is Nothing Then lastRow = fCell.Row
            
            Set fCell = .Cells.Find(What:="*", LookIn:=xlFormulas, _
                                   SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
            If Not fCell Is Nothing Then lastCol = fCell.Column
            
            Set fCell = .Cells.Find(What:="*", LookIn:=xlValues, _
                                   SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            If Not fCell Is Nothing Then lastRow = Application.Max(lastRow, fCell.Row)
            
            Set fCell = .Cells.Find(What:="*", LookIn:=xlValues, _
                                   SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
            If Not fCell Is Nothing Then lastCol = Application.Max(lastCol, fCell.Column)

            '--- Adjust for shapes extending range
            For Each shp In .Shapes
                On Error Resume Next
                shpLastRow = shp.BottomRightCell.Row
                shpLastCol = shp.BottomRightCell.Column
                ' Not On Error GoTo 0: that switched the whole procedure's
                ' handler off from the first shape onwards.
                On Error GoTo CleanUp
                
                lastRow = Application.Max(lastRow, shpLastRow)
                lastCol = Application.Max(lastCol, shpLastCol)
            Next shp

            '--- Prevent full sheet wipe
            If lastRow = 0 Then lastRow = 1
            If lastCol = 0 Then lastCol = 1

            '--- Delete unused columns
            If lastCol < .Columns.Count Then
                .Range(.Cells(1, lastCol + 1), .Cells(1, .Columns.Count)).EntireColumn.Delete
            End If

            '--- Delete unused rows
            If lastRow < .rows.Count Then
                .Range(.Cells(lastRow + 1, 1), .Cells(.rows.Count, 1)).EntireRow.Delete
            End If

        End With
    Next ws

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, "Shrink Excel File"

End Sub


' ----------------------------------------------------------------------------
'  Shared by ListFormulaToLeft and ListFormulaToLeftSelection.
'
'  Returns True when a listing was written.
'
'  Two things here are load bearing:
'
'   * .Formula, not .FormulaR1C1. The old version wrote "=""<--- "" &
'     FORMULATEXT(RC[-1])" through .FormulaR1C1. If the target cell carried a
'     Text number format Excel stored that string LITERALLY instead of parsing
'     it, and the R1C1 wording is what ended up on screen. Writing an A1
'     address through .Formula is what guarantees A1 style.
'
'   * The Text-format reset. Without it the same literal-string trap applies to
'     the new formula too.
'
'  Cells whose left neighbour holds no formula are left untouched rather than
'  filled with #N/A, which is all FORMULATEXT can return for a constant.
' ----------------------------------------------------------------------------
Private Function WriteFormulaListing(ByVal c As Range) As Boolean
    Dim srcCell As Range

    Set srcCell = c.Offset(0, -1)
    If Not srcCell.HasFormula Then Exit Function

    If c.NumberFormat = "@" Then c.NumberFormat = "General"

    c.formula = "=""<--- "" & FORMULATEXT(" & srcCell.Address(False, False) & ")"
    WriteFormulaListing = True
End Function


' ----------------------------------------------------------------------------
'  Shared close-out for the three Remove Formulas commands.
'
'  errNum / errTxt are captured by the caller at its CleanExit label BEFORE
'  FastModeOff runs. Reading Err in here instead would be reading it after two
'  further procedure calls, which is exactly the sort of thing that works until
'  one day it does not.
' ----------------------------------------------------------------------------
Private Sub ReportFormulaRemoval(ByVal Title As String, _
                                 ByVal cellsDone As Long, _
                                 ByVal areasFailed As Long, _
                                 ByVal skipped As String, _
                                 ByVal errNum As Long, _
                                 ByVal errTxt As String)
    Dim msg As String

    If errNum <> 0 Then
        MsgBox "Could not finish removing formulas." & vbCrLf & vbCrLf & _
               "Error " & errNum & ": " & errTxt, vbExclamation, Title
        Exit Sub
    End If

    If cellsDone = 0 And areasFailed = 0 And Len(skipped) = 0 Then
        MsgBox "No formulas were found.", vbInformation, Title
        Exit Sub
    End If

    msg = Format$(cellsDone, "#,##0") & " formula cell(s) replaced with values."

    If areasFailed > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              areasFailed & " block(s) were left alone. That normally means a " & _
              "multi-cell array formula, which Excel will not let you change a " & _
              "part of at a time."
    End If

    If Len(skipped) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Protected worksheets skipped:" & skipped
    End If

    MsgBox msg, vbInformation, Title
End Sub


' ----------------------------------------------------------------------------
'  Shared engine for the Wrap with ... commands.
'
'  Wraps the formula in every formula cell of the selection, skipping cells
'  that already start with skipIfStartsWith so running the tool twice does not
'  double up.
'
'  The six wrap / sign macros used to end with a hard
'  Application.Calculation = xlCalculationAutomatic, which yanked anyone
'  working in manual calc back to automatic, and none of them had an error
'  handler -- an error mid-loop left calculation stuck on manual for the rest
'  of the session. AppStateManager captures and restores the real prior state.
' ----------------------------------------------------------------------------
Private Sub WrapSelectionFormulas(ByVal Title As String, _
                                  ByVal wrapPrefix As String, _
                                  ByVal wrapSuffix As String, _
                                  ByVal skipIfStartsWith As String)
    Dim rg As Range
    Dim c As Range
    Dim n As Long
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(rg, False, Title) Then Exit Sub

    On Error GoTo CleanExit
    AppStateManager.FastModeOn

    For Each c In rg.Cells
        If c.HasFormula Then
            If Left$(c.formula, Len(skipIfStartsWith)) <> skipIfStartsWith Then
                ' Mid$(.., 2) drops the leading "=" of the existing formula.
                c.formula = wrapPrefix & Mid$(c.formula, 2) & wrapSuffix
                n = n + 1
            End If
        End If
    Next c

CleanExit:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    If errNum <> 0 Then
        MsgBox "Could not finish wrapping the selected formulas." & vbCrLf & vbCrLf & _
               "Error " & errNum & ": " & errTxt, vbExclamation, Title
    ElseIf n = 0 Then
        MsgBox "No formulas in the selection needed wrapping.", vbInformation, Title
    End If
End Sub

' ----------------------------------------------------------------------------
'  Shared engine for Change Sign of Inputs and Remove Negative Signs.
'
'  Only non-formula numeric cells are touched, which is what both prompts have
'  always promised. Remove Negative Signs used to open with
'  Selection.Value = Selection.Value, silently destroying every formula in the
'  selection before it started.
' ----------------------------------------------------------------------------
Private Sub ApplySignToInputCells(ByVal Title As String, ByVal makeAbsolute As Boolean)
    Dim rg As Range
    Dim c As Range
    Dim n As Long
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(rg, False, Title) Then Exit Sub

    On Error GoTo CleanExit
    AppStateManager.FastModeOn

    For Each c In rg.Cells
        If Not c.HasFormula Then
            If Not IsEmpty(c.Value2) Then
                ' IsNumeric(Empty) is True in VBA, hence the IsEmpty guard first.
                If IsNumeric(c.Value2) Then
                    If makeAbsolute Then
                        c.Value = Abs(c.Value2)
                    Else
                        c.Value = c.Value2 * -1
                    End If
                    n = n + 1
                End If
            End If
        End If
    Next c

CleanExit:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    If errNum <> 0 Then
        MsgBox "Could not finish." & vbCrLf & vbCrLf & _
               "Error " & errNum & ": " & errTxt, vbExclamation, Title
    ElseIf n = 0 Then
        MsgBox "No numeric input cells were found in the selection.", vbInformation, Title
    End If
End Sub


' ============================================================================
'  Page setup helpers
'
'  Three things changed across all six page-setup macros:
'
'  1. xSheet.Select is gone. Selecting a sheet individually BREAKS a multi-sheet
'     group -- so a macro whose entire purpose is "apply this to every selected
'     sheet" was destroying the selection it was working from, leaving only the
'     last sheet selected. PageSetup can be written without selecting anything,
'     so the grouping now survives and there is no screen flicker.
'
'  2. Application.PrintCommunication = False is deliberately NOT used here,
'     although it is the usual advice for making PageSetup faster. Measured on
'     this add-in it silently DROPS writes: with it off, clearing the three
'     footer zones left all three exactly as they were and raised no error, and
'     writing " " to all three applied only the last one. A single property per
'     sheet survived reliably, several did not. A page-setup macro that quietly
'     does nothing is far worse than a slow one.
'
'  3. These four lines are gone from Switch Portrait, Switch Landscape and
'     Std Page Setup:
'
'         With Application
'             .Calculation = xlAutomatic
'             .MaxChange = 0.001
'         End With
'         ActiveWorkbook.PrecisionAsDisplayed = False
'
'     They were macro-recorder leftovers. A button labelled "Switch Portrait"
'     has no business forcing a workbook back to automatic calculation, nor
'     touching Precision As Displayed, which is a data-affecting setting.
' ============================================================================

' Shared by Switch Portrait and Switch Landscape.
Private Sub ApplyPageOrientation(ByVal newOrientation As XlPageOrientation, _
                                 ByVal Title As String)
    Dim sh As Object
    Dim errNum As Long
    Dim errTxt As String

    If ActiveWindow Is Nothing Then Exit Sub

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    For Each sh In ActiveWindow.SelectedSheets
        If TypeOf sh Is Worksheet Then sh.PageSetup.Orientation = newOrientation
    Next sh

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, Title
End Sub
