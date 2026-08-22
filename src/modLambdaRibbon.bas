Attribute VB_Name = "modLambdaRibbon"
Option Explicit

' ============================================================================
'  modLambdaRibbon -- everything the LAMBDA Studio ribbon group talks to.
'
'  This is the TOP layer of three:
'
'     modLambdaRibbon   THIS FILE. Callbacks, selection state, every prompt.
'     modLambdaLib      moves LAMBDAs between places (no UI at all)
'     AddInStorage      the only thing that touches tblAlonzoChurch
'
'  The split is worth keeping. Everything that talks to the user lives here, so
'  modLambdaLib can be driven from the Immediate window without a dialog box
'  appearing -- which, with no test harness in this project, is how you actually
'  test it.
'
'  ---------------------------------------------------------------------------
'  TWO THINGS TO KNOW BEFORE EDITING
'  ---------------------------------------------------------------------------
'
'  1. EVERY CALLBACK MUST BE Public.
'     The ribbon calls these by name from outside the project. A Private one
'     produces a "callback not found" error at load with no clue as to which.
'
'  2. THE LISTS ARE CACHED ON PURPOSE. See mLibEntries / mAFEntries below.
'     Excel calls getItemLabel ONCE PER ITEM. The version this replaced re-read
'     the whole worksheet inside that callback, so opening a dropdown of 200
'     functions meant 200 full sheet scans. Now the list is built once, and
'     Lambda_Update throws the cache away.
' ============================================================================

Private Const DIALOG_TITLE As String = "XL Edge LAMBDA Library"
Private Const NONE_ITEM    As String = "[None]"

' The ribbon object, captured at load. See Lambda_OnLoad.
Public gRibbon As IRibbonUI

Private mSearchText     As String        ' contents of the filter box
Private mSelectedName   As String        ' highlighted item, library dropdown
Private mAFSelectedName As String        ' highlighted item, active-file dropdown

' Cached, filtered, sorted. Nothing = "not built yet, build on next request".
Private mLibEntries As Collection
Private mAFEntries  As Collection

' Which workbook mAFEntries was built from. The active-file list is only valid
' for one workbook, so the cache is keyed to it and rebuilds itself the moment a
' callback runs against a different one. Without this the list kept showing the
' functions from whichever workbook happened to be in front when Excel started.
Private mAFSource As String


' ============================================================================
'  RIBBON LIFECYCLE
' ============================================================================

' Called once by Excel when the add-in's ribbon loads. Wired up by the
' onLoad="Lambda_OnLoad" attribute on the <customUI> root element -- without
' that attribute this never runs, gRibbon stays Nothing, and no dropdown can
' ever refresh itself.
Public Sub Lambda_OnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

' Drop the cached lists and ask Excel to re-query the three controls.
'
' gRibbon is lost whenever VBA state resets -- an unhandled error, or simply
' editing code in the VBE. The Is Nothing guard means that degrades to "the
' dropdowns stop refreshing until Excel restarts" instead of raising error 91
' on every click from then on.
Public Sub Lambda_Update()
    Set mLibEntries = Nothing
    Set mAFEntries = Nothing
    mAFSource = vbNullString

    If gRibbon Is Nothing Then Exit Sub

    On Error Resume Next
    gRibbon.InvalidateControl "LambdaSearch"
    gRibbon.InvalidateControl "LambdaList"
    gRibbon.InvalidateControl "LambdaFunctions"
    On Error GoTo 0
End Sub

' Called from ThisWorkbook whenever a different workbook becomes active, one is
' opened, or one is created. See "ThisWorkbook (paste into the add-in).txt".
'
' WHY AN EVENT IS REQUIRED, and why "re-pick the item in the dropdown" cannot
' work instead:
'
'   * Excel does not re-query a ribbon control when you open the dropdown. It
'     serves whatever it cached last, until something calls InvalidateControl.
'   * A dropDown's onAction only fires when the SELECTION CHANGES. Re-picking
'     the item already highlighted -- [None], say -- fires nothing at all, so
'     there is no callback in which to notice the workbook changed.
'
' So nothing the user does inside the group can wake it up. The signal has to
' come from Excel itself, which is what the Application events provide.
Public Sub Lambda_ActiveFileChanged()
    On Error Resume Next
    Set mAFEntries = Nothing
    mAFSource = vbNullString
    mAFSelectedName = vbNullString
    If gRibbon Is Nothing Then Exit Sub
    gRibbon.InvalidateControl "LambdaFunctions"
End Sub


' ============================================================================
'  THE CACHED LISTS
' ============================================================================

' The library, filtered by the search box and sorted A-Z. Built at most once
' per Lambda_Update.
Private Function LibraryEntries() As Collection
    If Not mLibEntries Is Nothing Then
        Set LibraryEntries = mLibEntries
        Exit Function
    End If

    Dim all As Collection, hit As New Collection
    On Error Resume Next
    Set all = AddInStorage.LambdaAll()
    On Error GoTo 0
    If all Is Nothing Then Set all = New Collection

    Dim e As Variant
    For Each e In all
        ' Match on the NAME or the DESCRIPTION. (A Tags column used to be the
        ' second field searched; it was dropped because nobody filled it in.)
        If MatchesSearch(CStr(e(0)) & " " & CStr(e(2)), mSearchText) Then hit.Add e
    Next e

    Set mLibEntries = SortEntriesByName(hit)
    Set LibraryEntries = mLibEntries
End Function

' Every LAMBDA defined in the active workbook, sorted A-Z. Not filtered -- the
' search box belongs to the library dropdown.
'
' The cache is only reused while the ACTIVE WORKBOOK is still the one it was
' built from. Switch workbook and this rebuilds on the next callback, whether or
' not anything remembered to invalidate the ribbon.
Private Function AFEntries() As Collection
    Dim key As String
    key = ActiveWorkbookKey()

    If Not mAFEntries Is Nothing And key = mAFSource Then
        Set AFEntries = mAFEntries
        Exit Function
    End If

    ' Different workbook: the previous selection is meaningless now.
    If key <> mAFSource Then mAFSelectedName = ""

    Dim found As Collection
    On Error Resume Next
    If Application.Workbooks.Count > 0 Then
        Set found = modLambdaLib.EntriesFromWorkbookNames(ActiveWorkbook)
    End If
    On Error GoTo 0
    If found Is Nothing Then Set found = New Collection

    Set mAFEntries = SortEntriesByName(found)
    mAFSource = key
    Set AFEntries = mAFEntries
End Function

' Identifies the active workbook for cache purposes. FullName is empty for a
' workbook that has never been saved, so Name is included too -- otherwise every
' unsaved Book1/Book2 would look like the same workbook.
Private Function ActiveWorkbookKey() As String
    On Error Resume Next
    If Application.Workbooks.Count > 0 Then
        ActiveWorkbookKey = ActiveWorkbook.FullName & "|" & ActiveWorkbook.name
    End If
End Function

' True if EVERY space-separated term in search appears somewhere in haystack.
' An empty search matches everything. Terms are AND-ed, which is what makes
' typing "date month" narrow the list rather than widen it.
Private Function MatchesSearch(ByVal haystack As String, ByVal search As String) As Boolean
    Dim s As String
    s = Trim$(search)
    If Len(s) = 0 Then
        MatchesSearch = True
        Exit Function
    End If

    Dim terms() As String, i As Long
    terms = Split(s, " ")
    For i = LBound(terms) To UBound(terms)
        If Len(terms(i)) > 0 Then
            If InStr(1, haystack, terms(i), vbTextCompare) = 0 Then Exit Function
        End If
    Next i

    MatchesSearch = True
End Function

' Sort a Collection of Array(name, formula, description) by name, A-Z.
'
' Insertion sort: n is a few hundred at most, and it is close to free on an
' already-sorted list, which this usually is.
Private Function SortEntriesByName(ByVal entries As Collection) As Collection
    Dim n As Long
    n = entries.Count

    Dim out As New Collection
    Set SortEntriesByName = out
    If n = 0 Then Exit Function

    Dim buf() As Variant
    ReDim buf(1 To n)
    Dim i As Long
    For i = 1 To n
        buf(i) = entries(i)
    Next i

    Dim j As Long, key As Variant
    For i = 2 To n
        key = buf(i)
        j = i - 1
        Do While j >= 1
            If StrComp(CStr(buf(j)(0)), CStr(key(0)), vbTextCompare) > 0 Then
                buf(j + 1) = buf(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        buf(j + 1) = key
    Next i

    For i = 1 To n
        out.Add buf(i)
    Next i
End Function

' The entry at a ZERO-based ribbon index, or Empty if out of range.
' Ribbon indexes are 0-based; VBA Collections are 1-based. Every off-by-one bug
' in a dropdown callback starts here, so the conversion happens in exactly one
' place.
Private Function EntryAt(ByVal entries As Collection, ByVal zeroBasedIndex As Long) As Variant
    If zeroBasedIndex < 0 Then Exit Function
    If zeroBasedIndex + 1 > entries.Count Then Exit Function
    EntryAt = entries(zeroBasedIndex + 1)
End Function

' 0-based position of name in entries, or 0 if absent.
Private Function IndexOfName(ByVal entries As Collection, ByVal name As String) As Long
    Dim i As Long
    For i = 1 To entries.Count
        If StrComp(CStr(entries(i)(0)), name, vbTextCompare) = 0 Then
            IndexOfName = i - 1
            Exit Function
        End If
    Next i
End Function


' ============================================================================
'  CALLBACKS -- filter box  (editBox id="LambdaSearch")
' ============================================================================

Public Sub LambdaFilterChanged(control As IRibbonControl, text As String)
    mSearchText = text
    Lambda_Update
End Sub

Public Sub LambdaFilterText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = mSearchText
End Sub


' ============================================================================
'  CALLBACKS -- library dropdown  (dropDown id="LambdaList")
' ============================================================================

Public Sub LambdaListCount(control As IRibbonControl, ByRef returnedVal)
    Dim n As Long
    n = LibraryEntries().Count
    If n = 0 Then n = 1                          ' room for the [None] placeholder
    returnedVal = n
End Sub

Public Sub LambdaListLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    Dim e As Variant
    e = EntryAt(LibraryEntries(), CLng(index))
    If IsEmpty(e) Then returnedVal = NONE_ITEM Else returnedVal = e(0)
End Sub

' Which item starts out highlighted. Also the moment mSelectedName is reconciled
' with what is actually on the list -- after a delete or a filter change the
' previous selection may no longer exist.
Public Sub LambdaListSelected(control As IRibbonControl, ByRef returnedVal)
    Dim entries As Collection
    Set entries = LibraryEntries()

    Dim pos As Long
    pos = IndexOfName(entries, mSelectedName)
    returnedVal = pos

    Dim e As Variant
    e = EntryAt(entries, pos)
    If IsEmpty(e) Then mSelectedName = "" Else mSelectedName = CStr(e(0))
End Sub

Public Sub LambdaListScreentip(control As IRibbonControl, ByRef returnedVal)
    Dim d As String
    d = SelectedDescription()
    If Len(Trim$(d)) = 0 Then d = NONE_ITEM
    returnedVal = "Description: " & d
End Sub

' The supertip shows the start of the stored formula, so you can see what a
' function actually does before injecting it. (This slot used to show the Tags
' column, which no longer exists.)
Public Sub LambdaListSupertip(control As IRibbonControl, ByRef returnedVal)
    Const MAX_CHARS As Long = 250

    Dim f As String
    f = SelectedFormula()
    If Len(Trim$(f)) = 0 Then
        returnedVal = "Formula: " & NONE_ITEM
        Exit Sub
    End If

    ' Collapse in-cell line breaks -- a supertip renders them as literal boxes.
    f = Replace(Replace(Replace(f, vbCrLf, " "), vbLf, " "), vbCr, " ")
    Do While InStr(f, "  ") > 0
        f = Replace(f, "  ", " ")
    Loop

    If Len(f) > MAX_CHARS Then f = Left$(f, MAX_CHARS) & " ..."
    returnedVal = "Formula: " & f
End Sub

Public Sub LambdaListAction(control As IRibbonControl, id As String, index As Integer)
    Dim e As Variant
    e = EntryAt(LibraryEntries(), CLng(index))
    If IsEmpty(e) Then mSelectedName = "" Else mSelectedName = CStr(e(0))
    Lambda_Update
End Sub


' ============================================================================
'  CALLBACKS -- active file dropdown  (dropDown id="LambdaFunctions")
' ============================================================================

Public Sub LambdaAFCount(control As IRibbonControl, ByRef returnedVal)
    Dim n As Long
    n = AFEntries().Count
    If n = 0 Then n = 1
    returnedVal = n
End Sub

Public Sub LambdaAFLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    Dim e As Variant
    e = EntryAt(AFEntries(), CLng(index))
    If IsEmpty(e) Then returnedVal = NONE_ITEM Else returnedVal = e(0)
End Sub

Public Sub LambdaAFSelected(control As IRibbonControl, ByRef returnedVal)
    Dim entries As Collection
    Set entries = AFEntries()

    Dim pos As Long
    pos = IndexOfName(entries, mAFSelectedName)
    returnedVal = pos

    Dim e As Variant
    e = EntryAt(entries, pos)
    If IsEmpty(e) Then mAFSelectedName = "" Else mAFSelectedName = CStr(e(0))
End Sub

Public Sub LambdaAFAction(control As IRibbonControl, id As String, index As Integer)
    Dim e As Variant
    e = EntryAt(AFEntries(), CLng(index))
    If IsEmpty(e) Then mAFSelectedName = "" Else mAFSelectedName = CStr(e(0))
    Lambda_Update
End Sub


' ============================================================================
'  ACTIONS -- library to active file
' ============================================================================

Public Sub LambdaInjectSelected(control As IRibbonControl)
    On Error GoTo Oops
    If Not HaveWorkbook() Then Exit Sub
    If Not HaveSelection() Then Exit Sub

    Dim entries As Collection
    Set entries = modLambdaLib.EntriesFromLibrary(mSelectedName)
    If entries.Count = 0 Then
        Warn "'" & mSelectedName & "' is no longer in the library."
        Lambda_Update
        Exit Sub
    End If

    InjectAndReport entries, ActiveWorkbook
    Exit Sub
Oops:
    ShowError
End Sub

Public Sub LambdaInjectAll(control As IRibbonControl)
    On Error GoTo Oops
    If Not HaveWorkbook() Then Exit Sub

    Dim entries As Collection
    Set entries = modLambdaLib.EntriesFromLibrary()
    If entries.Count = 0 Then
        Warn "The LAMBDA library is empty."
        Exit Sub
    End If

    If MsgBox("Add all " & entries.Count & " library function(s) to '" & ActiveWorkbook.name & "'?" & _
              vbCrLf & vbCrLf & "Existing names with the same name will be replaced.", _
              vbQuestion + vbYesNo, DIALOG_TITLE) <> vbYes Then Exit Sub

    InjectAndReport entries, ActiveWorkbook
    Exit Sub
Oops:
    ShowError
End Sub

Public Sub LambdaInjectFromGist(control As IRibbonControl)
    On Error GoTo Oops
    If Not HaveWorkbook() Then Exit Sub

    Dim url As String
    url = AskGistUrl()
    If Len(url) = 0 Then Exit Sub

    Dim moduleName As String, cancelled As Boolean
    moduleName = AskModuleName(cancelled)
    If cancelled Then Exit Sub

    Dim entries As Collection
    Set entries = modLambdaLib.EntriesFromGistUrl(url)
    If entries.Count = 0 Then
        Warn "No LAMBDA / named-formula definitions found at that URL." & vbCrLf & _
             "Confirm it is a single-file Gist in the Advanced Formula Environment format."
        Exit Sub
    End If
    If Len(moduleName) > 0 Then Set entries = modLambdaLib.ApplyModuleNamespace(entries, moduleName)

    If MsgBox("Found " & entries.Count & " definition(s)." & vbCrLf & vbCrLf & _
              "Add them to '" & ActiveWorkbook.name & "'" & _
              IIf(Len(moduleName) > 0, " under module '" & moduleName & "'", "") & "?" & vbCrLf & _
              "Existing names with the same name will be replaced.", _
              vbQuestion + vbYesNo, DIALOG_TITLE) <> vbYes Then Exit Sub

    InjectAndReport entries, ActiveWorkbook
    Exit Sub
Oops:
    ShowError
End Sub

' Shared tail of the three inject actions: write the names, report the result.
'
' FastModeOn/Off are paired on a cleanup label, per AppStateManager's own
' guidance. That matters more than usual here: an add-in changes GLOBAL Excel
' state, so bailing out with EnableEvents still False would leave events dead
' for every open workbook until Excel restarts.
Private Sub InjectAndReport(ByVal entries As Collection, ByVal wb As Workbook)
    Dim failed As New Collection
    Dim wrapped As New Collection
    Dim added As Long

    On Error GoTo CleanExit
    AppStateManager.FastModeOn
    added = modLambdaLib.AddEntriesToWorkbook(entries, wb, failed, wrapped)

CleanExit:
    AppStateManager.FastModeOff
    If Err.Number <> 0 Then Err.Raise Err.Number, "InjectAndReport", Err.description

    Dim msg As String
    msg = "Added " & added & " of " & entries.Count & " function(s) to '" & wb.name & "'."

    ' Anything Excel would only accept wrapped is called differently, so say so.
    ' Silently changing how a function is called would be worse than the
    ' original failure -- the user would type Name, get an error, and have no
    ' way of knowing why.
    If wrapped.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "Excel would not store these as plain named formulas, so they were " & _
              "stored as zero-argument functions." & vbCrLf & _
              "Call them with brackets - NAME() rather than NAME:" & NameList(wrapped)
    End If

    If failed.Count > 0 Then msg = msg & vbCrLf & vbCrLf & _
                                   "Skipped (invalid name or conflict):" & NameList(failed)

    MsgBox msg, vbInformation, DIALOG_TITLE
End Sub


' ============================================================================
'  ACTIONS -- into the library
' ============================================================================

Public Sub LambdaImportSelectedAF(control As IRibbonControl)
    On Error GoTo Oops
    If Not HaveWorkbook() Then Exit Sub

    If Len(mAFSelectedName) = 0 Then
        Warn "Select a LAMBDA from the 'Active File (AF) LAMBDAs' list first."
        Exit Sub
    End If

    Dim entries As Collection
    Set entries = modLambdaLib.EntriesFromWorkbookNames(ActiveWorkbook, mAFSelectedName)
    If entries.Count = 0 Then
        Warn "'" & mAFSelectedName & "' is no longer defined in '" & ActiveWorkbook.name & "'."
        Lambda_Update
        Exit Sub
    End If

    StoreAndReport entries
    Exit Sub
Oops:
    ShowError
End Sub

Public Sub LambdaImportAllAF(control As IRibbonControl)
    On Error GoTo Oops
    If Not HaveWorkbook() Then Exit Sub

    Dim entries As Collection
    Set entries = modLambdaLib.EntriesFromWorkbookNames(ActiveWorkbook)
    If entries.Count = 0 Then
        Warn "'" & ActiveWorkbook.name & "' has no workbook-scoped LAMBDA functions."
        Exit Sub
    End If

    If MsgBox("Import all " & entries.Count & " LAMBDA(s) from '" & ActiveWorkbook.name & _
              "' into the library?", vbQuestion + vbYesNo, DIALOG_TITLE) <> vbYes Then Exit Sub

    StoreAndReport entries
    Exit Sub
Oops:
    ShowError
End Sub

Public Sub LambdaImportFromGist(control As IRibbonControl)
    On Error GoTo Oops

    Dim url As String
    url = AskGistUrl()
    If Len(url) = 0 Then Exit Sub

    Dim moduleName As String, cancelled As Boolean
    moduleName = AskModuleName(cancelled)
    If cancelled Then Exit Sub

    Dim entries As Collection
    Set entries = modLambdaLib.EntriesFromGistUrl(url)
    If entries.Count = 0 Then
        Warn "No LAMBDA / named-formula definitions found at that URL."
        Exit Sub
    End If
    If Len(moduleName) > 0 Then Set entries = modLambdaLib.ApplyModuleNamespace(entries, moduleName)

    If MsgBox("Found " & entries.Count & " definition(s). Add them to the library?", _
              vbQuestion + vbYesNo, DIALOG_TITLE) <> vbYes Then Exit Sub

    StoreAndReport entries
    Exit Sub
Oops:
    ShowError
End Sub

' Shared tail of every "into the library" action.
Private Sub StoreAndReport(ByVal entries As Collection)
    Dim added As Long, updated As Long
    Dim ok As Boolean

    On Error GoTo CleanExit
    AppStateManager.FastModeOn
    ok = modLambdaLib.AddEntriesToLibrary(entries, added, updated)

CleanExit:
    AppStateManager.FastModeOff
    If Err.Number <> 0 Then Err.Raise Err.Number, "StoreAndReport", Err.description

    If Not ok Then
        Warn "The library could not be updated." & vbCrLf & vbCrLf & AddInStorage.LastError
        Exit Sub
    End If

    ' Remember the last thing imported -- it is almost always what the user
    ' wants to act on next.
    If entries.Count > 0 Then mSelectedName = CStr(entries(entries.Count)(0))
    Lambda_Update

    MsgBox "Library updated." & vbCrLf & vbCrLf & _
           added & " added" & vbCrLf & _
           updated & " replaced", vbInformation, DIALOG_TITLE
End Sub


' ============================================================================
'  ACTIONS -- manage the library
' ============================================================================

Public Sub LambdaRename(control As IRibbonControl)
    On Error GoTo Oops
    If Not HaveSelection() Then Exit Sub

    Dim newName As String, cancelled As Boolean
    Do
        newName = Trim$(AskText("Rename '" & mSelectedName & "' to:", _
                                mSelectedName, cancelled))
        If cancelled Then Exit Sub

        If Len(newName) = 0 Then
            Warn "An empty string is not a valid name."
        ElseIf Not modLambdaLib.IsValidName(newName) Then
            Warn "'" & newName & "' is not a valid function name." & vbCrLf & vbCrLf & _
                 "Start with a letter or underscore, then letters, digits, " & _
                 "underscores or dots. No spaces."
        Else
            Exit Do
        End If
    Loop

    If Not AddInStorage.RenameLambda(mSelectedName, newName) Then
        Warn AddInStorage.LastError
        Exit Sub
    End If

    AddInStorage.SaveStorage
    mSelectedName = newName
    Lambda_Update
    Exit Sub
Oops:
    ShowError
End Sub

Public Sub LambdaDescribe(control As IRibbonControl)
    On Error GoTo Oops
    If Not HaveSelection() Then Exit Sub

    ' An empty answer here is legitimate -- it clears the description. Only a
    ' genuine Cancel (detected via StrPtr inside AskText) backs out.
    Dim answer As String, cancelled As Boolean
    answer = AskText("Description for '" & mSelectedName & "':", _
                     SelectedDescription(), cancelled)
    If cancelled Then Exit Sub

    If Not AddInStorage.SetLambdaDescription(mSelectedName, answer) Then
        Warn AddInStorage.LastError
        Exit Sub
    End If

    AddInStorage.SaveStorage
    Lambda_Update
    Exit Sub
Oops:
    ShowError
End Sub

Public Sub LambdaDelete(control As IRibbonControl)
    On Error GoTo Oops
    If Not HaveSelection() Then Exit Sub

    If MsgBox("'" & mSelectedName & "' will be deleted from the library." & vbCrLf & vbCrLf & _
              "This does not affect any workbook that already uses it. Continue?", _
              vbExclamation + vbYesNo + vbDefaultButton2, DIALOG_TITLE) <> vbYes Then Exit Sub

    If Not AddInStorage.DeleteLambda(mSelectedName) Then
        Warn AddInStorage.LastError
        Exit Sub
    End If

    AddInStorage.SaveStorage
    mSelectedName = ""
    Lambda_Update
    Exit Sub
Oops:
    ShowError
End Sub


' ============================================================================
'  ACTIONS -- export and import the whole library
'
'  Two formats on purpose:
'    .xlsx  a table you can read, edit and hand to anyone. The backup format.
'    .txt   Advanced Formula Environment / Gist format. Round-trips with Excel
'           Labs and diffs cleanly in source control. The power-user format.
'  Both read back through the same code path as a GitHub import.
' ============================================================================

Public Sub LambdaExportXlsx(control As IRibbonControl)
    On Error GoTo Oops
    If Not HaveLibrary() Then Exit Sub

    Dim path As Variant
    path = Application.GetSaveAsFilename( _
        InitialFileName:="XL Edge LAMBDA library.xlsx", _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
        Title:="Export LAMBDA library to Excel file")
    If VarType(path) = vbBoolean Then Exit Sub

    modLambdaLib.ExportLibraryToXlsx CStr(path)
    MsgBox AddInStorage.LambdaCount() & " function(s) exported to:" & vbCrLf & vbCrLf & path, _
           vbInformation, DIALOG_TITLE
    Exit Sub
Oops:
    ShowError
End Sub

Public Sub LambdaImportXlsx(control As IRibbonControl)
    On Error GoTo Oops

    Dim path As Variant
    path = Application.GetOpenFilename( _
        FileFilter:="Excel files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", _
        Title:="Import LAMBDA library from Excel file")
    If VarType(path) = vbBoolean Then Exit Sub

    Dim entries As Collection
    Set entries = modLambdaLib.EntriesFromXlsxFile(CStr(path))
    If entries.Count = 0 Then
        Warn "No functions found in that file." & vbCrLf & vbCrLf & _
             "It needs a 'Function' column and a 'Formula' column, with a " & _
             "'Description' column optional."
        Exit Sub
    End If

    If MsgBox("Found " & entries.Count & " function(s). Add them to the library?", _
              vbQuestion + vbYesNo, DIALOG_TITLE) <> vbYes Then Exit Sub

    StoreAndReport entries
    Exit Sub
Oops:
    ShowError
End Sub

Public Sub LambdaExportText(control As IRibbonControl)
    On Error GoTo Oops
    If Not HaveLibrary() Then Exit Sub

    Dim path As Variant
    path = Application.GetSaveAsFilename( _
        InitialFileName:="XL Edge LAMBDA library.txt", _
        FileFilter:="Text file (*.txt), *.txt", _
        Title:="Export LAMBDA library to text (Gist format)")
    If VarType(path) = vbBoolean Then Exit Sub

    modLambdaLib.ExportLibraryToTextFile CStr(path)
    MsgBox AddInStorage.LambdaCount() & " function(s) exported to:" & vbCrLf & vbCrLf & path, _
           vbInformation, DIALOG_TITLE
    Exit Sub
Oops:
    ShowError
End Sub

Public Sub LambdaImportText(control As IRibbonControl)
    On Error GoTo Oops

    Dim path As Variant
    path = Application.GetOpenFilename( _
        FileFilter:="Text files (*.txt), *.txt", _
        Title:="Import LAMBDA library from text (Gist format)")
    If VarType(path) = vbBoolean Then Exit Sub

    Dim entries As Collection
    Set entries = modLambdaLib.EntriesFromTextFile(CStr(path))
    If entries.Count = 0 Then
        Warn "No LAMBDA definitions found in that file." & vbCrLf & vbCrLf & _
             "Expected Advanced Formula Environment format: NAME = LAMBDA(...);"
        Exit Sub
    End If

    If MsgBox("Found " & entries.Count & " function(s). Add them to the library?", _
              vbQuestion + vbYesNo, DIALOG_TITLE) <> vbYes Then Exit Sub

    StoreAndReport entries
    Exit Sub
Oops:
    ShowError
End Sub


' ============================================================================
'  PROMPTS AND GUARDS
' ============================================================================

' Ask for a line of text.
'
' USE THIS, NOT Application.InputBox. Application.InputBox is Excel's
' CELL-PICKER dialog: it stays live on the grid so you can point at a range, so
' the arrow keys move the cell cursor and paste a reference into the box instead
' of moving the caret through the text. That is correct for "select a range" and
' completely wrong for "type a name" -- Type:=2 changes what it returns, not how
' it handles the keyboard.
'
' VBA's own InputBox is an ordinary modal dialog with ordinary text editing:
' arrows, Home/End and shift-select all behave as they should.
'
' The catch it has instead is that it returns "" for BOTH Cancel and an empty
' OK. StrPtr tells them apart -- on Cancel VBA hands back a null string pointer,
' on OK a real pointer to a zero-length string. That distinction is what lets
' the description prompt clear a description rather than treating it as Cancel.
Private Function AskText(ByVal prompt As String, ByVal defaultText As String, _
                         ByRef cancelled As Boolean) As String
    Dim answer As String
    answer = InputBox(prompt, DIALOG_TITLE, defaultText)

    cancelled = (StrPtr(answer) = 0)
    If Not cancelled Then AskText = answer
End Function

Private Function AskGistUrl() As String
    Dim cancelled As Boolean
    AskGistUrl = Trim$(AskText( _
        "GitHub Gist URL (page or raw):" & vbCrLf & vbCrLf & _
        "e.g. https://gist.github.com/jonwittwer/13e1c25374ef9de7d708e43db9e0f442", _
        "", cancelled))
End Function

' The optional Excel Labs "Module name". Sets cancelled so a blank answer (import
' as plain global names) can be told apart from pressing Cancel.
Private Function AskModuleName(ByRef cancelled As Boolean) As String
    Dim answer As String
    answer = AskText( _
        "Module name (optional) - like the Excel Labs field." & vbCrLf & _
        "e.g. VLL  ->  VLL.RESCALE, VLL.SFROUND ..." & vbCrLf & vbCrLf & _
        "References between the imported functions are rewritten to match." & vbCrLf & _
        "Leave BLANK to import plain global names.", _
        "", cancelled)
    If cancelled Then Exit Function

    Dim m As String
    m = Trim$(answer)
    If Len(m) > 0 Then
        If Not modLambdaLib.IsValidName(m) Then
            Warn "'" & m & "' is not a valid module name."
            cancelled = True
            Exit Function
        End If
    End If

    AskModuleName = m
End Function

Private Function HaveWorkbook() As Boolean
    If Application.Workbooks.Count = 0 Then
        Warn "Open a workbook first - there is nothing to act on."
        Exit Function
    End If
    HaveWorkbook = True
End Function

Private Function HaveSelection() As Boolean
    If Len(mSelectedName) = 0 Or mSelectedName = NONE_ITEM Then
        Warn "Select a function from the 'LAMBDA Library functions' list first."
        Exit Function
    End If
    HaveSelection = True
End Function

Private Function HaveLibrary() As Boolean
    If AddInStorage.LambdaCount() = 0 Then
        Warn "The LAMBDA library is empty - there is nothing to export."
        Exit Function
    End If
    HaveLibrary = True
End Function

Private Function SelectedDescription() As String
    If Len(mSelectedName) = 0 Then Exit Function
    On Error Resume Next
    SelectedDescription = AddInStorage.LambdaDescription(mSelectedName)
End Function

Private Function SelectedFormula() As String
    If Len(mSelectedName) = 0 Then Exit Function
    On Error Resume Next
    SelectedFormula = AddInStorage.LambdaFormula(mSelectedName)
End Function

' Bullet list of names, capped so a failed bulk import cannot produce a message
' box taller than the screen.
Private Function NameList(ByVal names As Collection) As String
    Const MAX_SHOWN As Long = 20
    Dim s As String, i As Long
    For i = 1 To names.Count
        If i > MAX_SHOWN Then
            s = s & vbCrLf & "  ...and " & (names.Count - MAX_SHOWN) & " more."
            Exit For
        End If
        s = s & vbCrLf & "  - " & names(i)
    Next i
    NameList = s
End Function

Private Sub Warn(ByVal message As String)
    MsgBox message, vbExclamation, DIALOG_TITLE
End Sub

' One place that turns a raised error into a message.
'
' The FastModeOff is a safety net, not the primary unwind -- InjectAndReport and
' StoreAndReport already balance their own. It is a no-op when the depth counter
' is already 0, and it is what stops an unexpected error path from leaving Excel
' with events off and the screen frozen, which reads to the user as a hang.
Private Sub ShowError()
    Dim n As Long, d As String
    n = Err.Number
    d = Err.description

    On Error Resume Next
    AppStateManager.FastModeOff
    On Error GoTo 0

    If n <> 0 Then MsgBox "Error " & n & ":" & vbCrLf & vbCrLf & d, vbExclamation, DIALOG_TITLE
End Sub
