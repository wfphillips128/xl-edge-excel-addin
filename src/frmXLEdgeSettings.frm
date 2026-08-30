VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmXLEdgeSettings 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmXLEdgeSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmXLEdgeSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
'  Controls, created at run time in UserForm_Initialize.
'
'  WithEvents is what lets a control created in code still fire Click events.
'  A control dragged on in the designer gets that wiring for free; one added
'  with Controls.Add does not, unless it is assigned to a WithEvents variable.
'  Controls we never need to react to (labels, text boxes) are declared plain.
' ----------------------------------------------------------------------------
Private mpgSettings   As MSForms.MultiPage
Private lblConstName  As MSForms.label
Private lblHint       As MSForms.label
Private txtValue      As MSForms.TextBox
Private txtCompany    As MSForms.TextBox

Private WithEvents lstConstants As MSForms.ListBox
Attribute lstConstants.VB_VarHelpID = -1
Private WithEvents cmdUpdate    As MSForms.CommandButton
Attribute cmdUpdate.VB_VarHelpID = -1
Private WithEvents lstCompanies As MSForms.ListBox
Attribute lstCompanies.VB_VarHelpID = -1
Private WithEvents cmdAdd       As MSForms.CommandButton
Attribute cmdAdd.VB_VarHelpID = -1
Private WithEvents cmdEdit      As MSForms.CommandButton
Attribute cmdEdit.VB_VarHelpID = -1
Private WithEvents cmdRemove    As MSForms.CommandButton
Attribute cmdRemove.VB_VarHelpID = -1
Private WithEvents cmdUp        As MSForms.CommandButton
Attribute cmdUp.VB_VarHelpID = -1
Private WithEvents cmdDown      As MSForms.CommandButton
Attribute cmdDown.VB_VarHelpID = -1
Private WithEvents cmdSave      As MSForms.CommandButton
Attribute cmdSave.VB_VarHelpID = -1
Private WithEvents cmdCancel    As MSForms.CommandButton
Attribute cmdCancel.VB_VarHelpID = -1

' ----------------------------------------------------------------------------
'  Working copies.
'
'  The form edits THESE, never the worksheet. Nothing touches the reference
'  sheet until Save is clicked -- which is what makes Cancel actually cancel.
' ----------------------------------------------------------------------------
Private mKeys()   As String
Private mVals()   As Variant
Private mCount    As Long
Private mCompanies As Collection
Private mSaved    As Boolean

' Value kinds, used for validation and for formatting the display.
Private Enum eValType
    vtNumber = 1
    vtDate = 2
    vtText = 3
End Enum

' ============================================================================
'  Setup
' ============================================================================

Private Sub UserForm_Initialize()
    BuildLayout
    LoadConstants
    LoadCompanies
    RefreshConstantList
    RefreshCompanyList
    ShowSelectedConstant
End Sub

Private Sub BuildLayout()
    Me.caption = "XL Edge Settings"
    Me.width = 486
    Me.Height = 412

    Set mpgSettings = Me.Controls.Add("Forms.MultiPage.1", "mpgSettings", True)
    With mpgSettings
        .Left = 6: .Top = 6: .width = 468: .Height = 336
        .Pages(0).caption = "Constants"
        .Pages(1).caption = "Companies"
    End With

    BuildConstantsPage mpgSettings.Pages(0)
    BuildCompaniesPage mpgSettings.Pages(1)

    Set cmdSave = Me.Controls.Add("Forms.CommandButton.1", "cmdSave", True)
    With cmdSave
        .caption = "Save": .Left = 296: .Top = 350: .width = 84: .Height = 24
        ' Deliberately NOT .Default. With a default button, pressing Enter after
        ' typing a value would save and close the form -- when what the user
        ' almost certainly meant was "apply this one edit". Save stays a
        ' deliberate click.
    End With

    Set cmdCancel = Me.Controls.Add("Forms.CommandButton.1", "cmdCancel", True)
    With cmdCancel
        .caption = "Cancel": .Left = 386: .Top = 350: .width = 84: .Height = 24
        .Cancel = True                     ' Esc closes the form
    End With
End Sub

Private Sub BuildConstantsPage(ByVal pg As Object)
    Dim c As Object

    Set c = pg.Controls.Add("Forms.Label.1", "lblListHdr", True)
    With c
        .caption = "Click a setting to edit its value:"
        .Left = 8: .Top = 6: .width = 400: .Height = 12
    End With

    Set lstConstants = pg.Controls.Add("Forms.ListBox.1", "lstConstants", True)
    With lstConstants
        .Left = 8: .Top = 20: .width = 440: .Height = 168
        .ColumnCount = 2
        .ColumnWidths = "190 pt;244 pt"
        .ColumnHeads = False
    End With

    Set lblConstName = pg.Controls.Add("Forms.Label.1", "lblConstName", True)
    With lblConstName
        .caption = ""
        .Left = 8: .Top = 194: .width = 440: .Height = 12
        .Font.bold = True
    End With

    Set txtValue = pg.Controls.Add("Forms.TextBox.1", "txtValue", True)
    With txtValue
        .Left = 8: .Top = 210: .width = 346: .Height = 34
        .Multiline = True                  ' the footer text is long
        .WordWrap = True
        .EnterKeyBehavior = False
    End With

    Set cmdUpdate = pg.Controls.Add("Forms.CommandButton.1", "cmdUpdate", True)
    With cmdUpdate
        .caption = "Update": .Left = 360: .Top = 210: .width = 88: .Height = 24
    End With

    Set lblHint = pg.Controls.Add("Forms.Label.1", "lblHint", True)
    With lblHint
        .caption = ""
        .Left = 8: .Top = 250: .width = 440: .Height = 52
        .WordWrap = True
    End With
End Sub

Private Sub BuildCompaniesPage(ByVal pg As Object)
    Dim c As Object

    Set c = pg.Controls.Add("Forms.Label.1", "lblCoHdr", True)
    With c
        .caption = "Names cycled by the Insert Company Name button:"
        .Left = 8: .Top = 6: .width = 400: .Height = 12
    End With

    Set lstCompanies = pg.Controls.Add("Forms.ListBox.1", "lstCompanies", True)
    With lstCompanies
        .Left = 8: .Top = 20: .width = 306: .Height = 216
    End With

    Set txtCompany = pg.Controls.Add("Forms.TextBox.1", "txtCompany", True)
    With txtCompany
        .Left = 8: .Top = 244: .width = 306: .Height = 18
    End With

    Set c = pg.Controls.Add("Forms.Label.1", "lblCoTip", True)
    With c
        .caption = "Type a name above, then Add. Select a row and use Rename or Remove. " & _
                   "Order here is the order the button cycles through."
        .Left = 8: .Top = 266: .width = 440: .Height = 30
        .WordWrap = True
    End With

    Set cmdAdd = pg.Controls.Add("Forms.CommandButton.1", "cmdAdd", True)
    With cmdAdd
        .caption = "Add": .Left = 324: .Top = 20: .width = 124: .Height = 24
    End With

    Set cmdEdit = pg.Controls.Add("Forms.CommandButton.1", "cmdEdit", True)
    With cmdEdit
        .caption = "Rename": .Left = 324: .Top = 48: .width = 124: .Height = 24
    End With

    Set cmdRemove = pg.Controls.Add("Forms.CommandButton.1", "cmdRemove", True)
    With cmdRemove
        .caption = "Remove": .Left = 324: .Top = 76: .width = 124: .Height = 24
    End With

    Set cmdUp = pg.Controls.Add("Forms.CommandButton.1", "cmdUp", True)
    With cmdUp
        .caption = "Move Up": .Left = 324: .Top = 112: .width = 124: .Height = 24
    End With

    Set cmdDown = pg.Controls.Add("Forms.CommandButton.1", "cmdDown", True)
    With cmdDown
        .caption = "Move Down": .Left = 324: .Top = 140: .width = 124: .Height = 24
    End With
End Sub

' ============================================================================
'  Per-setting metadata
'
'  Kept here rather than on the sheet so tblConstants stays two plain columns.
'  If you add a new setting, add it to both Select Case blocks -- an unlisted
'  key still works, it just gets treated as a number with no hint text.
' ============================================================================

Private Function KeyType(ByVal key As String) As eValType
    Select Case UCase$(Trim$(key))
        Case "PERIOD_END", "ITD_START"
            KeyType = vtDate
        Case "MYFOOTER"
            KeyType = vtText
        Case Else
            KeyType = vtNumber
    End Select
End Function

Private Function KeyHint(ByVal key As String) As String
    Select Case UCase$(Trim$(key))
        Case "PERIOD_END"
            KeyHint = "Last day of the reporting period. Drives the timeline slicers " & _
                      "and the month / quarter / year start dates."
        Case "ITD_START"
            KeyHint = "Inception-to-date start. The earliest date any ITD timeline filter reaches back to."
        Case "FORMULABAR_HEIGHT_SM"
            KeyHint = "Collapsed formula bar height, in lines. Toggled by Ctrl-Shift-U."
        Case "FORMULABAR_HEIGHT_LG"
            KeyHint = "Expanded formula bar height, in lines. Toggled by Ctrl-Shift-U."
        Case "LOWER_BOUND"
            KeyHint = "Smallest value produced by the random-fill macros."
        Case "UPPER_BOUND"
            KeyHint = "Largest value produced by the random-fill macros."
        Case "HEADER_ROW_HEIGHT_SINGLE"
            KeyHint = "Row height, in points, for a pivot header that occupies one row."
        Case "HEADER_ROW_HEIGHT_MULTI"
            KeyHint = "Row height, in points, for each row of a multi-row pivot header."
        Case "DEFAULT_ROW_HEIGHT"
            KeyHint = "Row height, in points, applied by the Set Row Heights buttons."
        Case "DEFAULT_COLUMN_WIDTH"
            KeyHint = "Column width, in characters, applied by the Set Column Widths buttons."
        Case "FOOTER_FONT_SIZE"
            KeyHint = "Point size of the page footer. Stored separately from the wording so " & _
                      "you never have to type Excel's '&' formatting codes."
        Case "MYFOOTER"
            KeyHint = "Confidentiality text printed in the page footer. Plain wording only -- " & _
                      "the point size is applied automatically from FOOTER_FONT_SIZE."
        Case Else
            KeyHint = "(No description available for this setting.)"
    End Select
End Function

Private Function TypeLabel(ByVal t As eValType) As String
    Select Case t
        Case vtDate:   TypeLabel = "date"
        Case vtText:   TypeLabel = "text"
        Case Else:     TypeLabel = "number"
    End Select
End Function

' Format a stored value for display.
Private Function DisplayValue(ByVal key As String, ByVal v As Variant) As String
    If IsEmpty(v) Then Exit Function
    Select Case KeyType(key)
        Case vtDate
            If IsDate(v) Then
                DisplayValue = Format$(CDate(v), "m/d/yyyy")
            Else
                DisplayValue = CStr(v)
            End If
        Case Else
            DisplayValue = CStr(v)
    End Select
End Function

' ============================================================================
'  Load / refresh
' ============================================================================

Private Sub LoadConstants()
    Dim keys As Variant
    keys = AddInStorage.ConstantKeys()

    mCount = 0
    ReDim mKeys(0 To 0)
    ReDim mVals(0 To 0)
    If Not AddInStorage.HasItems(keys) Then Exit Sub

    mCount = UBound(keys) - LBound(keys) + 1
    ReDim mKeys(1 To mCount)
    ReDim mVals(1 To mCount)

    Dim i As Long, src As Long
    src = LBound(keys)
    For i = 1 To mCount
        mKeys(i) = Trim$(CStr(keys(src)))
        mVals(i) = AddInStorage.ConstantValue(mKeys(i))
        src = src + 1
    Next i
End Sub

Private Sub LoadCompanies()
    Set mCompanies = New Collection

    Dim names As Variant
    names = AddInStorage.CompanyList()
    If Not AddInStorage.HasItems(names) Then Exit Sub

    Dim i As Long
    For i = LBound(names) To UBound(names)
        mCompanies.Add CStr(names(i))
    Next i
End Sub

Private Sub RefreshConstantList()
    Dim keep As Long
    keep = lstConstants.ListIndex

    lstConstants.Clear
    Dim i As Long
    For i = 1 To mCount
        lstConstants.AddItem mKeys(i)
        lstConstants.List(lstConstants.ListCount - 1, 1) = DisplayValue(mKeys(i), mVals(i))
    Next i

    If keep >= 0 And keep < lstConstants.ListCount Then
        lstConstants.ListIndex = keep
    ElseIf lstConstants.ListCount > 0 Then
        lstConstants.ListIndex = 0
    End If
End Sub

Private Sub RefreshCompanyList()
    Dim keep As Long
    keep = lstCompanies.ListIndex

    lstCompanies.Clear
    Dim i As Long
    For i = 1 To mCompanies.Count
        lstCompanies.AddItem mCompanies(i)
    Next i

    If keep >= 0 And keep < lstCompanies.ListCount Then
        lstCompanies.ListIndex = keep
    ElseIf lstCompanies.ListCount > 0 Then
        lstCompanies.ListIndex = 0
    End If
End Sub

Private Sub ShowSelectedConstant()
    Dim i As Long
    i = lstConstants.ListIndex + 1                 ' ListIndex is 0-based
    If i < 1 Or i > mCount Then
        lblConstName.caption = ""
        lblHint.caption = ""
        txtValue.text = ""
        txtValue.Enabled = False
        cmdUpdate.Enabled = False
        Exit Sub
    End If

    txtValue.Enabled = True
    cmdUpdate.Enabled = True
    lblConstName.caption = mKeys(i) & "   (" & TypeLabel(KeyType(mKeys(i))) & ")"
    lblHint.caption = KeyHint(mKeys(i))
    txtValue.text = DisplayValue(mKeys(i), mVals(i))
End Sub

' ============================================================================
'  Constants page events
' ============================================================================

Private Sub lstConstants_Click()
    ShowSelectedConstant
End Sub

Private Sub cmdUpdate_Click()
    Dim i As Long
    i = lstConstants.ListIndex + 1
    If i < 1 Or i > mCount Then Exit Sub

    Dim raw As String
    raw = Trim$(txtValue.text)

    ' Validate against the setting's kind BEFORE accepting it. Catching a bad
    ' entry here is the whole reason the form exists -- the same typo made
    ' directly on the sheet would silently fall back to the built-in default.
    Select Case KeyType(mKeys(i))

        Case vtDate
            If Not IsDate(raw) Then
                MsgBox mKeys(i) & " must be a date, for example 6/30/2026." & vbCrLf & vbCrLf & _
                       "'" & raw & "' was not recognised.", vbExclamation, "XL Edge Settings"
                txtValue.SetFocus
                Exit Sub
            End If
            mVals(i) = CDate(raw)

        Case vtNumber
            If Len(raw) = 0 Or Not IsNumeric(raw) Then
                MsgBox mKeys(i) & " must be a number." & vbCrLf & vbCrLf & _
                       "'" & raw & "' was not recognised.", vbExclamation, "XL Edge Settings"
                txtValue.SetFocus
                Exit Sub
            End If
            If CDbl(raw) < 0 Then
                MsgBox mKeys(i) & " cannot be negative.", vbExclamation, "XL Edge Settings"
                txtValue.SetFocus
                Exit Sub
            End If
            mVals(i) = CDbl(raw)

        Case Else
            If InStr(1, raw, "&") > 0 Then
                If MsgBox("'&' has a special meaning in an Excel footer -- it starts a " & _
                          "formatting code, so the text after it may not print as typed." & vbCrLf & vbCrLf & _
                          "Keep it anyway?", vbQuestion + vbYesNo, "XL Edge Settings") = vbNo Then
                    txtValue.SetFocus
                    Exit Sub
                End If
            End If
            mVals(i) = raw

    End Select

    RefreshConstantList
    ShowSelectedConstant
End Sub

' ============================================================================
'  Companies page events
' ============================================================================

Private Sub cmdAdd_Click()
    Dim s As String
    s = Trim$(txtCompany.text)
    If Len(s) = 0 Then
        MsgBox "Type a name in the box first.", vbInformation, "XL Edge Settings"
        txtCompany.SetFocus
        Exit Sub
    End If
    If CompanyExists(s) Then
        MsgBox "'" & s & "' is already in the list.", vbInformation, "XL Edge Settings"
        txtCompany.SetFocus
        Exit Sub
    End If

    mCompanies.Add s
    txtCompany.text = ""
    RefreshCompanyList
    lstCompanies.ListIndex = lstCompanies.ListCount - 1
    txtCompany.SetFocus
End Sub

Private Sub cmdEdit_Click()
    Dim i As Long
    i = lstCompanies.ListIndex + 1
    If i < 1 Then
        MsgBox "Select a name to rename.", vbInformation, "XL Edge Settings"
        Exit Sub
    End If

    Dim s As String
    s = Trim$(InputBox("New name:", "Rename Company", mCompanies(i)))
    If Len(s) = 0 Then Exit Sub                 ' cancelled or cleared
    If StrComp(s, mCompanies(i), vbTextCompare) <> 0 Then
        If CompanyExists(s) Then
            MsgBox "'" & s & "' is already in the list.", vbInformation, "XL Edge Settings"
            Exit Sub
        End If
    End If

    ' A Collection has no "replace" -- remove then insert at the same position
    ' so the cycling order is preserved.
    ReplaceCompany i, s
    RefreshCompanyList
    lstCompanies.ListIndex = i - 1
End Sub

Private Sub cmdRemove_Click()
    Dim i As Long
    i = lstCompanies.ListIndex + 1
    If i < 1 Then
        MsgBox "Select a name to remove.", vbInformation, "XL Edge Settings"
        Exit Sub
    End If

    If MsgBox("Remove '" & mCompanies(i) & "' from the list?", _
              vbQuestion + vbYesNo, "XL Edge Settings") = vbNo Then Exit Sub

    mCompanies.Remove i
    RefreshCompanyList
    If lstCompanies.ListCount > 0 Then
        If i - 1 < lstCompanies.ListCount Then
            lstCompanies.ListIndex = i - 1
        Else
            lstCompanies.ListIndex = lstCompanies.ListCount - 1
        End If
    End If
End Sub

Private Sub cmdUp_Click()
    MoveCompany -1
End Sub

Private Sub cmdDown_Click()
    MoveCompany 1
End Sub

Private Sub MoveCompany(ByVal delta As Long)
    Dim i As Long, j As Long
    i = lstCompanies.ListIndex + 1
    If i < 1 Then Exit Sub
    j = i + delta
    If j < 1 Or j > mCompanies.Count Then Exit Sub

    Dim s As String
    s = mCompanies(i)
    mCompanies.Remove i
    If j > mCompanies.Count Then
        mCompanies.Add s
    Else
        mCompanies.Add s, , j
    End If

    RefreshCompanyList
    lstCompanies.ListIndex = j - 1
End Sub

Private Sub ReplaceCompany(ByVal pos As Long, ByVal newName As String)
    mCompanies.Remove pos
    If pos > mCompanies.Count Then
        mCompanies.Add newName
    Else
        mCompanies.Add newName, , pos
    End If
End Sub

Private Function CompanyExists(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To mCompanies.Count
        If StrComp(mCompanies(i), s, vbTextCompare) = 0 Then
            CompanyExists = True
            Exit Function
        End If
    Next i
End Function

' ============================================================================
'  Save / Cancel
' ============================================================================

Private Sub cmdSave_Click()
    On Error GoTo Failed

    If Not AddInStorage.StorageReady() Then
        MsgBox "The add-in's reference sheet could not be found, so settings " & _
               "cannot be saved." & vbCrLf & vbCrLf & _
               "The add-in will keep running on its built-in defaults.", _
               vbExclamation, "XL Edge Settings"
        Exit Sub
    End If

    Dim failedKeys As String
    Dim i As Long
    For i = 1 To mCount
        If Not AddInStorage.SetConstant(mKeys(i), mVals(i)) Then
            failedKeys = failedKeys & vbCrLf & "   " & mKeys(i) & _
                         "  -  " & AddInStorage.LastError
        End If
    Next i

    ' Collection -> 1-based array for the storage layer.
    Dim names As Variant
    If mCompanies.Count > 0 Then
        ReDim names(1 To mCompanies.Count)
        For i = 1 To mCompanies.Count
            names(i) = mCompanies(i)
        Next i
    Else
        names = Array()
    End If

    If Not AddInStorage.SetCompanyList(names) Then
        MsgBox "The company list could not be written to the reference sheet." & vbCrLf & vbCrLf & _
               AddInStorage.LastError, vbExclamation, "XL Edge Settings"
        Exit Sub
    End If

    ' Persist the .XLAM itself. Without this the edits live only in memory and
    ' Excel discards them at shutdown without ever prompting.
    AddInStorage.SaveStorage
    AddInStorage.InvalidateCache

    If Len(failedKeys) > 0 Then
        MsgBox "Settings saved, but these keys were not found in tblConstants " & _
               "and were skipped:" & vbCrLf & failedKeys, vbExclamation, "XL Edge Settings"
    End If

    mSaved = True
    Me.Hide                     ' Hide, not Unload -- see the Saved property below
    Exit Sub

Failed:
    MsgBox "Could not save settings." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.description, vbExclamation, "XL Edge Settings"
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

' True if the user saved rather than cancelled.
'
' This is why the buttons call Me.Hide instead of Unload Me: unloading destroys
' the form's variables, so the calling macro would have nothing left to read.
' Hiding returns control to the caller with the object still intact; the caller
' unloads it when done.
Public Property Get Saved() As Boolean
    Saved = mSaved
End Property


