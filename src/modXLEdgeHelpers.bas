Attribute VB_Name = "modXLEdgeHelpers"
Option Explicit

' ============================================================================
'  modXLEdgeHelpers
'  Small, dependency-free helpers shared by the XL Edge ribbon macros.
'
'  Why a separate module: the same four guards (is there a selection, is the
'  range blank, is the user sure, turn formulas into values) were being
'  rewritten inline in every macro that needed them, each time slightly
'  differently. One copy means one place to fix.
'
'  Nothing in here touches the ribbon, a form, or AddInStorage, so it can be
'  imported into any project on its own.
' ============================================================================

' ----------------------------------------------------------------------------
'  Selection guards
' ----------------------------------------------------------------------------

' Standard opening line for a range macro:
'
'     If Not GetSelectionRange(rg, True, "My Tool") Then Exit Sub
'
' Returns False (and has already told the user why) when there is no range
' selected, or when singleAreaOnly is True and the user Ctrl-picked several
' blocks. rg is only Set when the function returns True.
Public Function GetSelectionRange(ByRef rg As Range, _
                                  Optional ByVal singleAreaOnly As Boolean = True, _
                                  Optional ByVal Title As String = "XL Edge") As Boolean

    Set rg = Nothing

    ' TypeName rather than TypeOf: TypeOf raises on a Nothing selection, which
    ' is exactly the case being guarded against.
    If TypeName(Application.Selection) <> "Range" Then
        MsgBox "Please select a range of cells first.", vbExclamation, Title
        Exit Function
    End If

    Set rg = Application.Selection

    If singleAreaOnly Then
        If rg.Areas.Count > 1 Then
            MsgBox "Please select a single block of cells." & vbCrLf & vbCrLf & _
                   "This tool works out what to do from the shape of the selection, " & _
                   "so it cannot run across several separate blocks.", _
                   vbExclamation, Title
            Set rg = Nothing
            Exit Function
        End If
    End If

    GetSelectionRange = True
End Function

' True when the range holds nothing at all.
'
' CountA counts a formula that returns "" as OCCUPIED. That is deliberate: a
' row of formulas showing blanks is not an empty row, and deleting it would
' throw away work.
Public Function IsBlankRange(ByVal rg As Range) As Boolean
    If rg Is Nothing Then
        IsBlankRange = True
    Else
        IsBlankRange = (Application.WorksheetFunction.CountA(rg) = 0)
    End If
End Function

' ----------------------------------------------------------------------------
'  Confirmation dialogs
' ----------------------------------------------------------------------------

' For anything that destroys data. Defaults to No, so a stray Enter is safe.
Public Function ConfirmDestructive(ByVal msg As String, ByVal Title As String) As Boolean
    ConfirmDestructive = (MsgBox(msg, vbYesNo Or vbExclamation Or vbDefaultButton2, Title) = vbYes)
End Function

' For anything reversible. Replaces the old "type Y or N" input boxes, which
' were case sensitive -- typing a lower case y silently did nothing.
Public Function ConfirmProceed(ByVal msg As String, ByVal Title As String) As Boolean
    ConfirmProceed = (MsgBox(msg, vbYesNo Or vbQuestion, Title) = vbYes)
End Function

' ----------------------------------------------------------------------------
'  Formulas to values
' ----------------------------------------------------------------------------

' Replaces every formula in rg with its current value.
'
' Note what is NOT used here: Range.HasFormula. On a range holding both
' formulas and constants it returns Null, and "If Null Then" is False -- so
' "If rg.HasFormula Then rg.Value = rg.Value" quietly does nothing on exactly
' the sheets people run it on. SpecialCells has no such trap.
'
' areasFailed counts blocks that refused to convert. In practice that means a
' multi-cell array (CSE) formula, where Excel will not let part of an array be
' changed. Those are skipped rather than aborting the whole run.
Public Sub ConvertFormulasToValues(ByVal rg As Range, _
                                   ByRef cellsDone As Long, _
                                   ByRef areasFailed As Long)
    Dim fx As Range
    Dim ar As Range

    If rg Is Nothing Then Exit Sub

    On Error Resume Next
    Set fx = rg.SpecialCells(xlCellTypeFormulas)
    Err.Clear
    On Error GoTo 0

    If fx Is Nothing Then Exit Sub

    For Each ar In fx.Areas
        On Error Resume Next
        Err.Clear
        ar.Value = ar.Value
        If Err.Number <> 0 Then
            areasFailed = areasFailed + 1
            Err.Clear
        Else
            cellsDone = cellsDone + ar.Cells.Count
        End If
        On Error GoTo 0
    Next ar
End Sub

' ----------------------------------------------------------------------------
'  Add-in plumbing
' ----------------------------------------------------------------------------

' OnKey and OnTime resolve a bare procedure name against the ACTIVE workbook,
' not the add-in. From an .XLAM that means the name is never found and the
' shortcut silently does nothing. Qualifying it with the add-in's own file
' name is what makes it resolve.
Public Function QualifiedMacro(ByVal procName As String) As String
    QualifiedMacro = "'" & ThisWorkbook.name & "'!" & procName
End Function

' ----------------------------------------------------------------------------
'  Error reporting
' ----------------------------------------------------------------------------

' The standard tail for a macro that has entered fast mode:
'
'     On Error GoTo CleanUp
'     AppStateManager.FastModeOn
'     ...work...
' CleanUp:
'     errNum = Err.Number
'     errTxt = Err.description
'     AppStateManager.FastModeOff
'     ReportError errNum, errTxt, TITLE_TXT
'
' Err must be read into locals BEFORE FastModeOff runs -- reading it afterwards
' means reading it across an intervening procedure call, which works right up
' until the day something in that call touches Err.
'
' Does nothing when errNum is 0, so the same line serves the success path.
Public Sub ReportError(ByVal errNum As Long, ByVal errTxt As String, ByVal Title As String)
    If errNum = 0 Then Exit Sub
    MsgBox Title & " could not finish." & vbCrLf & vbCrLf & _
           "Error " & errNum & ": " & errTxt, vbExclamation, Title
End Sub

' Trims a string the way the "Trim Spaces" button promises: leading and
' trailing spaces removed AND runs of spaces inside the text collapsed to one.
'
' VBA's own Trim() only does the ends. WorksheetFunction.Trim is the one that
' collapses internal runs.
'
' Chr(160) -- the non-breaking space -- is converted to an ordinary space
' first. It arrives with anything pasted from a web page or a PDF, is
' indistinguishable on screen, and defeats every trim there is. It is the usual
' reason two account names that look identical refuse to match.
Public Function CleanSpaces(ByVal s As String) As String
    If InStr(1, s, Chr$(160)) > 0 Then s = Replace(s, Chr$(160), " ")
    CleanSpaces = Application.WorksheetFunction.Trim(s)
End Function
