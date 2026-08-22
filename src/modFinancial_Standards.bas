Attribute VB_Name = "modFinancial_Standards"
Option Explicit

' HEADER_ROW_HEIGHT_SINGLE / _MULTI and FORMULABAR_HEIGHT_SM / _LG moved to
' tblConstants on the add-in's reference sheet. They are still visible here as
' global Property Get procedures in AddInStorage, so every reference below is
' unchanged. Edit them from the XL Edge settings form, not in code.



Public Sub CreateShortcuts()
'Registers the XL Edge keyboard shortcuts.
'   Ctrl = ^   Shift = +   Alt = %
'
'Two things were wrong with the old version, and together they meant NONE of
'the shortcuts advertised in the ribbon supertips worked in the shipped .XLAM:
'
' 1. It was called from Auto_Open. Auto_Open does not fire for a workbook that
'    Excel loads as an add-in -- Workbook_Open does. ThisWorkbook now calls
'    this directly. (Auto_Open still calls it too, for when the .xlsm itself is
'    opened during development.)
'
' 2. The macro names were unqualified. A bare name is resolved against the
'    ACTIVE workbook, not the add-in, so from an .XLAM it is never found.
'    QualifiedMacro prefixes the add-in's own file name, the same trick the
'    OnTime call in modSettingsUI already uses.
'
'The old version also did Application.OnKey "{F1}", "" -- disabling Excel Help
'for the whole session, never restoring it, and doing so on behalf of a user
'who only ticked a formatting add-in. That has been dropped.

    On Error Resume Next
    Application.OnKey "+^1", QualifiedMacro("ToggleNumberFormats2")   'Ctrl-Shift-1
    Application.OnKey "+^2", QualifiedMacro("ToggleNumberScale2")     'Ctrl-Shift-2
    Application.OnKey "+^3", QualifiedMacro("ToggleFontSize2")        'Ctrl-Shift-3
    Application.OnKey "+^4", QualifiedMacro("ToggleFonts2")           'Ctrl-Shift-4
    Application.OnKey "+^6", QualifiedMacro("ToggleFontColors2")      'Ctrl-Shift-6
    Application.OnKey "+^7", QualifiedMacro("ToggleFillColors2")      'Ctrl-Shift-7
    Application.OnKey "+^8", QualifiedMacro("ToggleIndents2")         'Ctrl-Shift-8
    Application.OnKey "+^U", QualifiedMacro("ToggleFormulaBar2")      'Ctrl-Shift-U
End Sub


Public Sub DeleteShortcuts()
'Hands every key back to Excel. Called when the add-in unloads, so XL Edge does
'not leave keys hijacked in a session that is still running.

    On Error Resume Next
    Application.OnKey "+^1"
    Application.OnKey "+^2"
    Application.OnKey "+^3"
    Application.OnKey "+^4"
    Application.OnKey "+^6"
    Application.OnKey "+^7"
    Application.OnKey "+^8"
    Application.OnKey "+^U"
    Application.OnKey "{F1}"      'restores Help if an older build disabled it
End Sub


Private Sub Auto_Open()
'Only fires when this file is opened as an ordinary workbook -- i.e. while
'developing in XL Edge.xlsm. Once built to .XLAM, ThisWorkbook.Workbook_Open
'is what registers the shortcuts.

    modFinancial_Standards.CreateShortcuts
End Sub


Sub Ribbon_OpenVBAEditor(control As IRibbonControl)
    OpenVBAEditor2
End Sub


Sub OpenVBAEditor2()
'Opens the VBA editor.
'
'ExecuteMso is tried first because it is the same command as the Visual Basic
'button on the Developer tab and does NOT need "Trust access to the VBA project
'object model". Application.VBE does need it, and raises 1004 without it, so it
'is only the fallback.

    Dim ok As Boolean

    On Error Resume Next

    Application.CommandBars.ExecuteMso "VisualBasic"
    ok = (Err.Number = 0)
    Err.Clear

    If Not ok Then
        Application.VBE.MainWindow.Visible = True
        ok = (Err.Number = 0)
        Err.Clear
    End If

    On Error GoTo 0

    If Not ok Then
        MsgBox "Excel would not open the VBA editor." & vbCrLf & vbCrLf & _
               "Tick File > Options > Trust Center > Trust Center Settings > " & _
               "Macro Settings > ""Trust access to the VBA project object model""," & _
               vbCrLf & "or just press Alt+F11.", _
               vbExclamation, "Show VBA Editor"
    End If
End Sub


Sub ShowMacroDialog(control As IRibbonControl)
    Application.Dialogs(xlDialogRun).Show
End Sub


Sub ToggleFormulaBar(control As IRibbonControl)
    ToggleFormulaBar2
End Sub

Sub ToggleFormulaBar2()
    If Application.FormulaBarHeight = FORMULABAR_HEIGHT_SM Then
        Application.FormulaBarHeight = FORMULABAR_HEIGHT_LG
    Else
        Application.FormulaBarHeight = FORMULABAR_HEIGHT_SM
    End If
End Sub


Public Sub FormatPTDataFields(control As IRibbonControl, Optional ByVal DoRefresh As Boolean = False)
    Dim errNum As Long
    Dim errTxt As String

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim df As PivotField
    Dim hdr As Range
    Dim headerRows As Long
    Dim targetHeight As Double

    On Error GoTo CleanUp
    AppStateManager.FastModeOn

    For Each ws In ActiveWorkbook.Worksheets
        For Each pt In ws.PivotTables

            On Error GoTo PivotErr

            With pt
                .ManualUpdate = True

                .HasAutoFormat = False
                .PreserveFormatting = True
                .DisplayFieldCaptions = True
                .ShowDrillIndicators = False

                If DoRefresh Then .RefreshTable

                '=========== FORCE FULL LAYOUT INVALIDATION ==========
                .RowAxisLayout xlOutlineRow
                .RowAxisLayout xlTabularRow

                '=========== HEADER RANGE ===========================
                Set hdr = GetPivotHeaderRange(pt)
                If Not hdr Is Nothing Then

                    headerRows = hdr.rows.Count

                    If headerRows = 1 Then
                        targetHeight = HEADER_ROW_HEIGHT_SINGLE
                    Else
                        targetHeight = HEADER_ROW_HEIGHT_MULTI
                    End If

                    With hdr
                        .WrapText = True
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .EntireRow.rowHeight = targetHeight
                    End With

                End If

                '=========== DATA FIELD FORMATS =====================
                For Each df In .DataFields
                    ApplyDataFieldFormat df
                Next df

                .ManualUpdate = False
                .Update

                '=========== RESTORE ALIGNMENTS =====================
                NormalizePivotAlignments pt

            End With

PivotNext:
            On Error GoTo 0
            Set hdr = Nothing
            GoTo ContinueLoop

PivotErr:
            On Error Resume Next
            pt.ManualUpdate = False
            On Error GoTo 0
            Resume PivotNext

ContinueLoop:
        Next pt
    Next ws

CleanUp:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    ReportError errNum, errTxt, "Format Pivot Table"

End Sub


Private Function GetPivotHeaderRange(ByVal pt As PivotTable) As Range

    Dim tr As Range
    Dim hdrRows As Long

    On Error Resume Next
    Set tr = pt.TableRange1
    On Error GoTo 0
    If tr Is Nothing Then Exit Function

    If Not pt.DataBodyRange Is Nothing Then
        hdrRows = pt.DataBodyRange.Row - tr.Row
        If hdrRows < 1 Then hdrRows = 1
        Set GetPivotHeaderRange = tr.rows(1).Resize(hdrRows)
    Else
        Set GetPivotHeaderRange = tr.rows(1)
    End If

End Function


Private Sub NormalizePivotAlignments(ByVal pt As PivotTable)

    On Error Resume Next

    If Not pt.RowRange Is Nothing Then
        pt.RowRange.HorizontalAlignment = xlLeft
    End If

    If Not pt.DataBodyRange Is Nothing Then
        pt.DataBodyRange.HorizontalAlignment = xlRight
    End If

    On Error GoTo 0

End Sub


Private Sub ApplyDataFieldFormat(ByVal pvtFld As PivotField)

    Dim cap As String
    cap = pvtFld.Caption

    Select Case True
        Case InStr(1, cap, "Date", vbTextCompare) > 0
            pvtFld.NumberFormat = "_mm/dd/yyyy_;_(@_)"

        Case InStr(1, cap, "Var %", vbTextCompare) > 0 _
          Or InStr(1, cap, "Variance %", vbTextCompare) > 0 _
          Or InStr(1, cap, "Var Pct", vbTextCompare) > 0 _
          Or InStr(1, cap, "Variance Pct", vbTextCompare) > 0
            pvtFld.NumberFormat = "_(#,##0.0%_);(#,##0.0%);_(""-""_);_(@_)"

        Case InStr(1, cap, "Var", vbTextCompare) > 0
            pvtFld.NumberFormat = "[Color10]" & ChrW(&H25B2) & _
                                  "_(#,##0_);[Red]" & ChrW(&H25BC) & _
                                  "(#,##0);_(""-""_);_(@_)"

        Case InStr(1, cap, "%", vbTextCompare) > 0 _
          Or InStr(1, cap, "Pct", vbTextCompare) > 0 _
          Or InStr(1, cap, "Percent", vbTextCompare) > 0 _
          Or InStr(1, cap, "Ratio", vbTextCompare) > 0 _
          Or InStr(1, cap, "Rate", vbTextCompare) > 0 _
          Or InStr(1, cap, "CAGR", vbTextCompare) > 0 _
          Or InStr(1, cap, "Weights", vbTextCompare) > 0
            pvtFld.NumberFormat = "_(#,##0.0%_);(#,##0.0%);_(""-""_);_(@_)"

        Case InStr(1, cap, "Factor", vbTextCompare) > 0
            pvtFld.NumberFormat = "_(###0.000_);(###0.000);_(""-""_);_(@_)"

        Case InStr(1, cap, "Year", vbTextCompare) > 0
            pvtFld.NumberFormat = "_(###0_);(###0);_(""-""_);_(@_)"

        Case Else
            pvtFld.NumberFormat = "_(#,##0_);(#,##0);_(""-""_);_(@_)"
    End Select

End Sub


Sub ToggleNumberFormats(control As IRibbonControl)
    ToggleNumberFormats2
End Sub


Sub ToggleNumberFormats2()
'Assign to your preferred shortcut keys using the Application.OnKey method

    With Selection
    Select Case Selection.NumberFormat
        Case "_(###0_);(###0);_(""–""_);_(@_)"
            Selection.NumberFormat = "_(#,##0_);(#,##0);_(""–""_);_(@_)"
        Case "_(#,##0_);(#,##0);_(""–""_);_(@_)"
            Selection.NumberFormat = "_($#,##0_);($#,##0);_(""–""_);_(@_)"
        Case "_($#,##0_);($#,##0);_(""–""_);_(@_)"
            Selection.NumberFormat = "_(#,##0.0%_);(#,##0.0%);_(#,##0.0%_);_(@_)"
        Case "_(#,##0.0%_);(#,##0.0%);_(#,##0.0%_);_(@_)"
            Selection.NumberFormat = "_(0x_);(0x);_(""–""_);_(@_)"
        Case "_(0x_);(0x);_(""–""_);_(@_)"
            'below is the Variance format with up and down arrows assigned to positive and negative numbers
            Selection.NumberFormat = "[Color10]" & ChrW(&H25B2) & "_(#,##0_);[Red]" & ChrW(&H25BC) & "(#,##0);_(""–""_);_(@_)"
        Case "[Color10]" & ChrW(&H25B2) & "_(#,##0_);[Red]" & ChrW(&H25BC) & "(#,##0);_(""–""_);_(@_)"
        'below is the Variance % format with up and down arrows assigned to positive and negative numbers
            Selection.NumberFormat = "[Color10]" & ChrW(&H25B2) & "_(#,##0.0%_);[Red]" & ChrW(&H25BC) & "(#,##0.0%);_(#,##0.0%_);_(@_)"
        Case "[Color10]" & ChrW(&H25B2) & "_(#,##0.0%_);[Red]" & ChrW(&H25BC) & "(#,##0.0%);_(#,##0.0%_);_(@_)"
            Selection.NumberFormat = "_(###0.0000_);(###0.0000);_(""–""_);_(@_)"
        Case Else
            Selection.NumberFormat = "_(#,##0_);(#,##0);_(""–""_);_(@_)"
    End Select
    End With
    
End Sub

Sub ToggleNumberScale(control As IRibbonControl)
    ToggleNumberScale2
End Sub

Sub ToggleNumberScale2()

    With Selection
    
    
    Select Case Selection.NumberFormat
        Case "_(#,##0_);(#,##0);_(""–""_);_(@_)"
            Selection.NumberFormat = "_(#,##0.0,K_);(#,##0.0,K);_(""–""_);_(@_)"
        Case "_(#,##0.0,K_);(#,##0.0,K);_(""–""_);_(@_)"
            Selection.NumberFormat = "_(#,##0.0,,""M""_);(#,##0.0,,""M"");_(""–""_);_(@_)"
        Case "_(#,##0.0,,""M""_);(#,##0.0,,""M"");_(""–""_);_(@_)"
            Selection.NumberFormat = "_(#,##0.0,,,""B""_);(#,##0.0,,,""B"");_(""–""_);_(@_)"
        Case "_(#,##0.0,,,""B""_);(#,##0.0,,,""B"");_(""–""_);_(@_)"
            'below is the Variance format with up and down arrows assigned to positive and negative numbers
            Selection.NumberFormat = "[Color10]" & ChrW(&H25B2) & "_(#,##0_);[Red]" & ChrW(&H25BC) & "(#,##0);_(""–""_);_(@_)"
        Case Else
            Selection.NumberFormat = "_(#,##0_);(#,##0);_(""–""_);_(@_)"
    End Select
    End With
    
End Sub

Sub ToggleDateFormats(control As IRibbonControl)

    With Selection
    Select Case Selection.NumberFormat
        Case "mm/dd/yyyy;@"
            Selection.NumberFormat = "m/d/yyyy;@"
        Case "m/d/yyyy;@"
            Selection.NumberFormat = "m/d/yy;@"
        Case "m/d/yy;@"
            Selection.NumberFormat = "mm-dd-yyyy;@"
        Case "mm-dd-yyyy;@"
            Selection.NumberFormat = "m-d-yyyy;@"
        Case "m-d-yyyy;@"
            Selection.NumberFormat = "m-d-yy;@"
        Case "m-d-yy;@"
            Selection.NumberFormat = "0000"
        Case Else
            Selection.NumberFormat = "mm/dd/yyyy;@"
    End Select
    End With


End Sub

Sub BinaryFormatCycle(control As IRibbonControl)
'No short-cut assigned since use isn't expected to be common

    With Selection
    Select Case Selection.NumberFormat
        Case "_(#,##0_);(#,##0);_(""–""_);_(@_)"
            Selection.NumberFormat = "[=1]""Yes"";[=0]""No"""
        Case "[=1]""Yes"";[=0]""No"""
            Selection.NumberFormat = "[=1]""True"";[=0]""False"""
        Case "[=1]""True"";[=0]""False"""
            Selection.NumberFormat = "[=1]""On"";[=0]""Off"""
        Case "[=1]""On"";[=0]""Off"""
            Selection.NumberFormat = "[=1]""Pass"";[=0]""Fail"""
        Case Else 'Reset to standard numerical formating
            Selection.NumberFormat = "_(#,##0_);(#,##0);_(""–""_);_(@_)"
    End Select
    End With
    
End Sub


Sub FillRight(control As IRibbonControl)
'Copies the selected formulas to the right, one column at a time.
'
'A column is filled only when BOTH hold:
'   * every destination cell in that column is empty, and
'   * the cell directly ABOVE the top of the block, or directly BELOW the
'     bottom of it, holds a formula -- the guide row.
'The walk stops at the first column that fails either test.
'
'The previous version used Selection.Offset(-1, 0).End(xlToRight), which jumps
'to the end of the contiguous run in the row above. That is normally far to the
'right of where the fill should stop, and AutoFill then wrote over whatever was
'in the way.

    Const TITLE_TXT As String = "Fill Formulas to Right"

    Dim src As Range
    Dim ws As Worksheet
    Dim dest As Range
    Dim firstCol As Long
    Dim lastCol As Long
    Dim c As Long
    Dim topRow As Long
    Dim botRow As Long
    Dim filled As Long
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(src, True, TITLE_TXT) Then Exit Sub

    Set ws = src.Worksheet
    topRow = src.Row
    botRow = src.Row + src.rows.Count - 1
    firstCol = src.Column + src.Columns.Count
    lastCol = ws.Columns.Count

    For c = firstCol To lastCol
        Set dest = ws.Range(ws.Cells(topRow, c), ws.Cells(botRow, c))
        If Not IsBlankRange(dest) Then Exit For
        If Not HasFormulaAboveOrBelow(ws, topRow, botRow, c) Then Exit For
        filled = filled + 1
    Next c

    If filled = 0 Then
        MsgBox "Nothing to fill." & vbCrLf & vbCrLf & _
               "The next column to the right is either already occupied, or the " & _
               "rows directly above and below it hold no formula to follow.", _
               vbInformation, TITLE_TXT
        Exit Sub
    End If

    On Error GoTo CleanExit
    AppStateManager.FastModeOn

    ' AutoFill is safe here: every destination cell was checked as empty above.
    src.AutoFill Destination:=ws.Range(ws.Cells(topRow, src.Column), _
                                       ws.Cells(botRow, firstCol + filled - 1)), _
                 Type:=xlFillDefault

CleanExit:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    If errNum <> 0 Then
        MsgBox "Could not fill to the right." & vbCrLf & vbCrLf & _
               "Error " & errNum & ": " & errTxt, vbExclamation, TITLE_TXT
    End If
End Sub

' True when this column has a formula directly above the top of the block, or
' directly below the bottom of it. Row 1 and the last row are guarded so the
' lookup never falls off the sheet.
Private Function HasFormulaAboveOrBelow(ByVal ws As Worksheet, _
                                        ByVal topRow As Long, _
                                        ByVal botRow As Long, _
                                        ByVal col As Long) As Boolean
    If topRow > 1 Then
        If ws.Cells(topRow - 1, col).HasFormula Then
            HasFormulaAboveOrBelow = True
            Exit Function
        End If
    End If

    If botRow < ws.rows.Count Then
        If ws.Cells(botRow + 1, col).HasFormula Then HasFormulaAboveOrBelow = True
    End If
End Function


Sub FillDown(control As IRibbonControl)
'Fills the selected formulas downwards into blank rows only.
'
'The fill stops entirely at the first row where ANY destination cell is not
'blank, so nothing is ever overwritten. It is also bounded by the worksheet's
'used range, so a selection with nothing at all beneath it does not try to fill
'a million rows.
'
'The previous version used Selection.Offset(0, -1).End(xlDown) to pick the
'extent, which overshoots, and AutoFill then wrote over existing values.

    Const TITLE_TXT As String = "Fill Formula Down"

    Dim src As Range
    Dim ws As Worksheet
    Dim dest As Range
    Dim firstRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim leftCol As Long
    Dim rightCol As Long
    Dim filled As Long
    Dim errNum As Long
    Dim errTxt As String

    If Not GetSelectionRange(src, True, TITLE_TXT) Then Exit Sub

    Set ws = src.Worksheet
    leftCol = src.Column
    rightCol = src.Column + src.Columns.Count - 1
    firstRow = src.Row + src.rows.Count
    lastRow = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1

    For r = firstRow To lastRow
        Set dest = ws.Range(ws.Cells(r, leftCol), ws.Cells(r, rightCol))
        If Not IsBlankRange(dest) Then Exit For
        filled = filled + 1
    Next r

    If filled = 0 Then
        MsgBox "Nothing to fill." & vbCrLf & vbCrLf & _
               "The row directly below the selection is already occupied, or " & _
               "there is no data below the selection to fill alongside.", _
               vbInformation, TITLE_TXT
        Exit Sub
    End If

    On Error GoTo CleanExit
    AppStateManager.FastModeOn

    src.AutoFill Destination:=ws.Range(ws.Cells(src.Row, leftCol), _
                                       ws.Cells(firstRow + filled - 1, rightCol)), _
                 Type:=xlFillDefault

CleanExit:
    errNum = Err.Number
    errTxt = Err.description
    AppStateManager.FastModeOff
    If errNum <> 0 Then
        MsgBox "Could not fill down." & vbCrLf & vbCrLf & _
               "Error " & errNum & ": " & errTxt, vbExclamation, TITLE_TXT
    End If
End Sub



Sub AlignCycle(control As IRibbonControl)

With Selection
    Select Case Selection.HorizontalAlignment
        Case xlGeneral
            Selection.HorizontalAlignment = xlLeft
        Case xlLeft
            Selection.HorizontalAlignment = xlCenter
        Case xlCenter
            Selection.HorizontalAlignment = xlRight
        Case xlRight
            Selection.HorizontalAlignment = xlGeneral
        Case Else
            'Selection.HorizontalAlignment = xlGeneral
    End Select
End With
             

End Sub


Sub ToggleGrid(control As IRibbonControl)

    If ActiveWindow.DisplayGridlines = True Then
        ActiveWindow.DisplayGridlines = False
    Else
        ActiveWindow.DisplayGridlines = True
    End If
        
End Sub

Sub ToggleFontSize(control As IRibbonControl)
    ToggleFontSize2
End Sub


Sub ToggleFontSize2()

    With Selection
    Select Case Selection.Font.Size
        Case 14
            Selection.Font.Size = 8
        Case 8
            Selection.Font.Size = 9
        Case 9
            Selection.Font.Size = 10
        Case 10
            Selection.Font.Size = 11
        Case 11
            Selection.Font.Size = 14
        Case Else
            Selection.Font.Size = 8
    End Select
    End With
    
End Sub


Sub ToggleFonts(control As IRibbonControl)
    ToggleFonts2
End Sub


Sub ToggleFonts2()

On Error Resume Next

    With Selection
    Select Case Selection.Font.name
        Case "Times New Roman"
            Selection.Font.name = "Calibri"
        Case "Calibri"
            Selection.Font.name = "Arial"
        Case "Arial"
            Selection.Font.name = "Aptos"
        Case "Aptos"
            Selection.Font.name = "Open Sans"
        Case "Open Sans"
            Selection.Font.name = "Neue Haas Grotesk Text Pro"
        Case Else
            Selection.Font.name = "Times New Roman"
    End Select
    End With
    
End Sub


Sub ToggleIndents(control As IRibbonControl)
    ToggleIndents2
End Sub

Sub ToggleIndents2()

    With Selection
    Select Case Selection.IndentLevel
        Case 0
            Selection.IndentLevel = 1
        Case 1
            Selection.IndentLevel = 2
        Case 2
            Selection.IndentLevel = 3
        Case 3
            Selection.IndentLevel = 4
        Case 4
            Selection.IndentLevel = 5
        Case Else
            Selection.IndentLevel = 0
    End Select
    End With
    
End Sub


Sub ToggleFontsMaster(control As IRibbonControl)

On Error Resume Next

    With ActiveSheet.UsedRange
    Select Case ActiveSheet.UsedRange.Font.name
        Case "Times New Roman"
            ActiveSheet.UsedRange.Font.name = "Calibri"
        Case "Calibri"
            ActiveSheet.UsedRange.Font.name = "Arial"
        Case "Arial"
            ActiveSheet.UsedRange.Font.name = "Aptos"
        Case "Aptos"
            ActiveSheet.UsedRange.Font.name = "Open Sans"
        Case "Open Sans"
            ActiveSheet.UsedRange.Font.name = "Neue Haas Grotesk Text Pro"
        Case Else
            ActiveSheet.UsedRange.Font.name = "Times New Roman"
    End Select
    End With
    
End Sub


Sub ToggleFontColors(control As IRibbonControl)
    ToggleFontColors2
End Sub

Sub ToggleFontColors2()

Dim strInputs As Long
Dim strWorksheet As Long
Dim strFormulas As Long
Dim strPartials As Long
Dim strWorkbook As Long
Dim strFileLinks As Long
Dim strInvestigate As Long
Dim strHeading As Long
Dim strAltText As Long

strInputs = RGB(Red:=0, Green:=0, Blue:=204) 'Blue
strWorksheet = RGB(Red:=40, Green:=154, Blue:=114) 'Green
strFormulas = RGB(Red:=0, Green:=0, Blue:=0) 'Black
strPartials = RGB(Red:=250, Green:=98, Blue:=28) 'Orange
strWorkbook = RGB(Red:=121, Green:=34, Blue:=155) 'Purple
strFileLinks = RGB(Red:=195, Green:=40, Blue:=56) 'Red
strInvestigate = RGB(Red:=255, Green:=255, Blue:=153) 'Light Yellow
strHeading = RGB(Red:=19, Green:=46, Blue:=87) 'Dark Blue
strAltText = RGB(Red:=255, Green:=255, Blue:=255) 'White

    With Selection
        Select Case Selection.Font.Color
            Case strAltText
                Selection.Font.Color = strInputs
            Case strInputs
                Selection.Font.Color = strWorksheet
            Case strWorksheet
                Selection.Font.Color = strFormulas
            Case strFormulas
                Selection.Font.Color = strPartials
            Case strPartials
                Selection.Font.Color = strWorkbook
            Case strWorkbook
                Selection.Font.Color = strFileLinks
            Case strFileLinks
                Selection.Font.Color = strInvestigate
            Case strInvestigate
                Selection.Font.Color = strHeading
            Case strHeading
                Selection.Font.Color = strAltText
            Case Else
                Selection.Font.Color = strFormulas
        End Select
    End With

End Sub


Sub ToggleFillColors(control As IRibbonControl)
    ToggleFillColors2
End Sub


Sub ToggleFillColors2()

Dim strInputs As Long
Dim strWorksheet As Long
Dim strFormulas As Long
Dim strPartials As Long
Dim strWorkbook As Long
Dim strFileLinks As Long
Dim strInvestigate As Long
Dim strHeading As Long
Dim strHeading2 As Long
Dim strAltText As Long
Dim strLightShading As Long

strInputs = RGB(Red:=0, Green:=0, Blue:=204) 'Blue
strWorksheet = RGB(Red:=40, Green:=154, Blue:=114) 'Green
strFormulas = RGB(Red:=0, Green:=0, Blue:=0) 'Black
strPartials = RGB(Red:=250, Green:=98, Blue:=28) 'Orange
strWorkbook = RGB(Red:=121, Green:=34, Blue:=155) 'Purple
strFileLinks = RGB(Red:=195, Green:=40, Blue:=56) 'Red
strInvestigate = RGB(Red:=255, Green:=255, Blue:=153) 'Light Yellow
strHeading = RGB(Red:=19, Green:=46, Blue:=87) 'Dark Blue
strHeading2 = RGB(Red:=217, Green:=229, Blue:=247) 'Light Blue
strAltText = RGB(Red:=255, Green:=255, Blue:=255) 'White
strLightShading = RGB(Red:=242, Green:=242, Blue:=242) 'Light Grey

    With Selection
        Select Case Selection.Interior.Color
            Case strHeading2
                Selection.Interior.Color = strInputs
                Selection.Font.Color = strAltText
                Selection.Font.Bold = True
            Case strInputs
                Selection.Interior.Color = strWorksheet
                Selection.Font.Color = strAltText
                Selection.Font.Bold = True
            Case strWorksheet
                Selection.Interior.Color = strFormulas
                Selection.Font.Color = strAltText
                Selection.Font.Bold = True
            Case strFormulas
                Selection.Interior.Color = strPartials
                Selection.Font.Color = strAltText
                Selection.Font.Bold = True
            Case strPartials
                Selection.Interior.Color = strWorkbook
                Selection.Font.Color = strAltText
                Selection.Font.Bold = True
            Case strWorkbook
                Selection.Interior.Color = strFileLinks
                Selection.Font.Color = strAltText
                Selection.Font.Bold = True
            Case strFileLinks
                Selection.Interior.Color = strInvestigate
                Selection.Font.Color = strFileLinks
            Case strInvestigate
                Selection.Interior.Color = strLightShading
                Selection.Font.Color = strFormulas
                Selection.Font.Bold = True
            Case strLightShading
                Selection.Interior.Color = strHeading
                Selection.Font.Color = strAltText
                Selection.Font.Bold = True
            Case strHeading
                Selection.Interior.Color = strHeading2
                Selection.Font.Color = strInputs
                Selection.Font.Bold = True
            Case Else
                Selection.Interior.Color = strHeading2
                Selection.Font.Color = strInputs
                Selection.Font.Bold = True
        End Select
    End With

End Sub


Sub RemoveSpecialFormatting(control As IRibbonControl)

Dim strBlack As Long
Dim strOrange As Long
Dim strBlue As Long
Dim strWhite As Long
Dim fontchoice As String

strBlack = RGB(Red:=0, Green:=0, Blue:=0) 'Black
strOrange = RGB(Red:=250, Green:=98, Blue:=28) 'Orange
strBlue = RGB(Red:=19, Green:=46, Blue:=87) 'Dark Blue
strWhite = RGB(Red:=19, Green:=46, Blue:=87) 'White
fontchoice = "Calibri"
    'common choices to paste as desired: "Calibri", "Arial", "Aptos", "Open Sans", "Neue Haas Grotesk Text Pro", "Times New Roman", "Courier New"

    Selection.Borders.LineStyle = xlNone
    Selection.Font.Bold = False
    Selection.Interior.Color = xlNone
    Selection.Font.Color = vbBlack
    'Selection.Font.Size = 10
    'Selection.Font = fontchoice

End Sub


Sub RemoveBorders(control As IRibbonControl)

    Selection.Borders.LineStyle = xlNone

End Sub


Sub FinancialFontColorCoding(control As IRibbonControl)
'Uses different font color for cells having formulas, text, numbers and constants
'Shading does not include file links or references to other worksheets
'Run this code to change font colors in Used Range on Active Sheet

Dim cell As Range
Dim formulaColor As Long
Dim constantColor As Long
Dim formulawithtextColor As Long
Dim formulawithlinkColor As Long
Dim formulawithrefColor As Long
Dim formulawithotherColor As Long

formulaColor = RGB(Red:=0, Green:=0, Blue:=0) 'Black
constantColor = RGB(Red:=0, Green:=0, Blue:=204) 'Blue
formulawithtextColor = RGB(Red:=250, Green:=98, Blue:=28) 'Orange
formulawithlinkColor = RGB(Red:=121, Green:=34, Blue:=155) 'Purple
formulawithrefColor = RGB(Red:=40, Green:=154, Blue:=114) 'Green
formulawithotherColor = RGB(Red:=195, Green:=40, Blue:=55) 'Red

On Error Resume Next

    'color cells having formulas
    For Each cell In ActiveSheet.UsedRange.SpecialCells(xlCellTypeFormulas, xlTextValues)
    cell.Font.Color = formulawithtextColor
    Next cell
    
    'color cells having formulas and having nos.
    For Each cell In ActiveSheet.UsedRange.SpecialCells(xlCellTypeFormulas, xlNumbers)
    cell.Font.Color = formulaColor
    Next cell
    
    'color cells having constants with text (non-formulas)
    For Each cell In ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants, xlTextValues)
    cell.Font.Color = formulaColor
    Next cell
    
    'color cells having constants with numbers (non-formulas)
    For Each cell In ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers)
    cell.Font.Color = constantColor
    Next cell
    
    HighlightConstants ' referencee seperate macro that is more complicated which finds numbers hardcoded into formulas

End Sub

Sub AutoColorSelection(control As IRibbonControl)
'Range must be selected
'Will highlight a few additinal items versus the above macro


Dim cell As Range
Dim formulaColor As Long
Dim constantColor As Long
Dim formulawithtextColor As Long
Dim formulawithlinkColor As Long
Dim formulawithrefColor As Long
Dim formulawithotherColor As Long

formulaColor = RGB(Red:=0, Green:=0, Blue:=0) 'Black
constantColor = RGB(Red:=0, Green:=0, Blue:=204) 'Blue
formulawithtextColor = RGB(Red:=250, Green:=98, Blue:=28) 'Orange
formulawithlinkColor = RGB(Red:=121, Green:=34, Blue:=155) 'Purple
formulawithrefColor = RGB(Red:=40, Green:=154, Blue:=114) 'Green
formulawithotherColor = RGB(Red:=195, Green:=40, Blue:=55) 'Red

On Error Resume Next

    Selection.SpecialCells(xlCellTypeConstants, xlTextValues).Font.Color = formulaColor
    Selection.SpecialCells(xlCellTypeConstants, xlNumbers).Font.Color = constantColor
    Selection.SpecialCells(xlCellTypeFormulas, xlTextValues).Font.Color = formulawithtextColor
    Selection.SpecialCells(xlCellTypeFormulas, xlNumbers).Font.Color = formulaColor

    'to be more specific
    For Each cell In Selection
        If Left(cell.formula & " ", 1) = "=" Then
            If InStr(CleanStr(cell.formula), "]") Then
                cell.Font.Color = formulawithlinkColor
            ElseIf InStr(CleanStr(cell.formula), "!") Then
                cell.Font.Color = formulawithrefColor
            End If
        'ElseIf HasConstant( = True Then cell.Font.Color = constantColor
        End If
    Next cell

    HighlightConstants ' referencee seperate macro that is more complicated which finds numbers hardcoded into formulas
    
End Sub


Sub AutoBlueBlackSelection(control As IRibbonControl)
'Range must be selected
'Will highlight a few additinal items versus the above macro


Dim cell As Range
Dim formulaColor As Long
Dim constantColor As Long
Dim formulawithtextColor As Long
Dim formulawithlinkColor As Long
Dim formulawithrefColor As Long
Dim formulawithotherColor As Long

formulaColor = RGB(Red:=0, Green:=0, Blue:=0) 'Black
constantColor = RGB(Red:=0, Green:=0, Blue:=204) 'Blue
formulawithtextColor = RGB(Red:=250, Green:=98, Blue:=28) 'Orange
formulawithlinkColor = RGB(Red:=121, Green:=34, Blue:=155) 'Purple
formulawithrefColor = RGB(Red:=40, Green:=154, Blue:=114) 'Green
formulawithotherColor = RGB(Red:=195, Green:=40, Blue:=55) 'Red

On Error Resume Next

    Selection.SpecialCells(xlCellTypeConstants, xlTextValues).Font.Color = formulaColor
    Selection.SpecialCells(xlCellTypeConstants, xlNumbers).Font.Color = constantColor
    Selection.SpecialCells(xlCellTypeFormulas, xlTextValues).Font.Color = constantColor
    Selection.SpecialCells(xlCellTypeFormulas, xlNumbers).Font.Color = formulaColor
   
End Sub

Sub FinancialFontColorCoding2(control As IRibbonControl)
'Uses different font color for cells having formulas, text, numbers and constants
'Shading includes file links or references to other worksheets
'Run this code to change font colors in Used Range on Active Sheet

Dim cell As Range
Dim formulaColor As Long
Dim constantColor As Long
Dim formulawithtextColor As Long
Dim formulawithlinkColor As Long
Dim formulawithrefColor As Long
Dim formulawithotherColor As Long

formulaColor = RGB(Red:=0, Green:=0, Blue:=0) 'Black
constantColor = RGB(Red:=0, Green:=0, Blue:=204) 'Blue
formulawithtextColor = RGB(Red:=250, Green:=98, Blue:=28) 'Orange
formulawithlinkColor = RGB(Red:=121, Green:=34, Blue:=155) 'Purple
formulawithrefColor = RGB(Red:=40, Green:=154, Blue:=114) 'Green
formulawithotherColor = RGB(Red:=195, Green:=40, Blue:=55) 'Red

On Error Resume Next

    'color cells having formulas
    For Each cell In ActiveSheet.UsedRange.SpecialCells(xlCellTypeFormulas, xlTextValues)
    cell.Font.Color = formulawithtextColor
    Next cell
    
    'color cells having formulas and having nos.
    For Each cell In ActiveSheet.UsedRange.SpecialCells(xlCellTypeFormulas, xlNumbers)
    cell.Font.Color = formulaColor
    Next cell
    
    'color cells having constants with text (non-formulas)
    For Each cell In ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants, xlTextValues)
    cell.Font.Color = formulaColor
    Next cell
    
    'color cells having constants with numbers (non-formulas)
    For Each cell In ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers)
    cell.Font.Color = constantColor
    Next cell

    'to be more specifiy
    For Each cell In ActiveSheet.UsedRange
        If Left(cell.formula & " ", 1) = "=" Then
            If InStr(CleanStr(cell.formula), "]") Then
                cell.Font.Color = formulawithlinkColor
            ElseIf InStr(CleanStr(cell.formula), "!") Then
                cell.Font.Color = formulawithrefColor
            End If
        'ElseIf HasConstant( = True Then cell.Font.Color = constantColor
        End If
    Next cell

    HighlightConstants ' referencee seperate macro that is more complicated which finds numbers hardcoded into formulas
    
End Sub

Function CleanStr(strIn As String) As String

Dim objRegex As Object
Set objRegex = CreateObject("vbscript.regexp")
    
    With objRegex
       .Pattern = "\""[^)]*\"""
       .Global = True
       CleanStr = .Replace(strIn, vbNullString)
    End With

End Function


Sub HighlightConstants()
'The most difficult part of the standard Finance formats
'Uses the currently selected worksheet

Dim wb As Workbook
Dim sh As Worksheet
Dim rng As Range
Dim rng2 As Range
Dim rCell As Range
Dim aCell As Range
Dim arr As Variant
Dim i As Long
Dim sStr As String
Dim strName As String
Dim iCtr As Long
Dim formulawithtextColor As Long

Set wb = ActiveWorkbook
Set sh = wb.ActiveSheet
Set rng = sh.UsedRange
formulawithtextColor = RGB(Red:=250, Green:=98, Blue:=28)

On Error Resume Next '\\ In case no formulas
Set rng = rng.SpecialCells(xlFormulas)
On Error GoTo 0

arr = Array("/", "~*", "+", "-", ">", "<", "=", "^", "[*]", "(")

If Not rng Is Nothing Then
        For Each rCell In rng.Cells
        For i = LBound(arr) To UBound(arr)
        sStr = "*" & arr(i) & "[0-9]*"
            If rCell.formula Like sStr Then
                If Not rng2 Is Nothing Then
                Set rng2 = Union(rng2, rCell)
                Else
                Set rng2 = rCell
                End If
            End If
        Next i
        Next rCell
Else
'\\No formulas found
End If

If Not rng2 Is Nothing Then rng2.Font.Color = formulawithtextColor


End Sub


Sub ConstantsInFormulas()
'Macro runns on entire worksheet to identify hardcoded figures in formulas
'In additiona to highlighting the cells, it adds an additonal worksheet that calls out constants in formulas

Dim wb As Workbook
Dim sh As Worksheet
Dim rng As Range
Dim rng2 As Range
Dim rCell As Range
Dim aCell As Range
Dim arr As Variant
Dim sStr As String
Dim strName As String
Dim msg As String
Dim i As Long
Dim iCtr As Long
Dim formulawithtextColor As Long

'constantColor = RGB(Red:=0, Green:=0, Blue:=205)
formulawithtextColor = RGB(Red:=250, Green:=98, Blue:=28)
Set wb = ActiveWorkbook
Set sh = wb.ActiveSheet
Set rng = sh.UsedRange

On Error Resume Next '\\ In case no formulas!
Set rng = rng.SpecialCells(xlFormulas)
On Error GoTo 0

arr = Array("/", "~*", "+", "-", ">", "<", "=", "^", "[*]", "(")

If Not rng Is Nothing Then
For Each rCell In rng.Cells
For i = LBound(arr) To UBound(arr)
sStr = "*" & arr(i) & "[0-9]*"
If rCell.formula Like sStr Then
If Not rng2 Is Nothing Then
Set rng2 = Union(rng2, rCell)
Else
Set rng2 = rCell
End If
End If
Next i
Next rCell
Else
'\\No formulas found
End If

If Not rng2 Is Nothing Then
'\\ do something e.g.:
Debug.Print rng2.Address
'\\ Highlight Formulas with constants such as:
    'Rng2.Font.Color = formulawithtextColor

'\\ Add a report sheet
Sheets.Add
'\\ Name the report sheet -include Report date & time
strName = "FormulasReport" _
& Format(Now, "yyyymmdd hh-mm")
ActiveSheet.name = strName

For Each aCell In rng2.Cells
iCtr = iCtr + 1
'\\ Write information to the Report sheet
With ActiveSheet
.Cells(iCtr, "A") = aCell.Address(External:=True)
.Cells(iCtr, "B") = "'" & aCell.formula
End With
Next aCell

ActiveSheet.Columns("A:B").AutoFit

'\\ Parse address string to produce columnar MsgBox report
'\\ N.B. A Msgbox is limited to 255 characters.
msg = "Cells holding formulas which include constants" _
& vbNewLine
msg = msg & Replace(rng2.Address(0, 0), ",", Chr(10))

Else
msg = "No Formula constants found in " & sh.name
End If

MsgBox prompt:=msg, _
Buttons:=vbInformation, _
Title:="Formulas Report"

End Sub


