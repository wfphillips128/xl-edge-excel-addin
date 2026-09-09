Attribute VB_Name = "modLambdaLib"
Option Explicit

' ============================================================================
'  modLambdaLib -- the LAMBDA library ENGINE.
'
'  Three modules make the LAMBDA Studio menu work, and they stack:
'
'     modLambdaRibbon   ribbon callbacks, selection state, all MsgBox prompts
'     modLambdaLib      THIS FILE. Moves LAMBDAs between places.
'     AddInStorage      the only thing that touches tblAlonzoChurch
'
'  ---------------------------------------------------------------------------
'  THE ONE IDEA IN THIS FILE
'  ---------------------------------------------------------------------------
'  Eight of the menu items are the same two operations with different ends
'  plugged in. So everything is normalised to a single shape:
'
'       Collection of Array(name, formula, description)
'
'  and then it is only ever a matter of picking a source and a sink:
'
'    SOURCES                             SINKS
'      EntriesFromGistUrl        ---+
'      EntriesFromTextFile       ---+--> AddEntriesToLibrary   -> tblAlonzoChurch
'      EntriesFromXlsxFile       ---+
'      EntriesFromWorkbookNames  ---+
'      EntriesFromLibrary        ------> AddEntriesToWorkbook  -> Name Manager
'
'  "Import LAMBDAs from GitHub to library" is EntriesFromGistUrl into
'  AddEntriesToLibrary. "Inject entire LAMBDA library to active file" is
'  EntriesFromLibrary into AddEntriesToWorkbook. Adding a new source later --
'  a different site, a database -- means writing one function, not a new menu.
'
'  The formula in an entry ALWAYS carries its leading "=", exactly as a defined
'  name's RefersTo does. The Gist text format is the one place it doesn't, and
'  that is handled at the edges (ProcessEntry adds it, FormatAsGistText strips
'  it) so nothing in the middle has to think about it.
'
'  Scope: no project references required. WinHTTP, Scripting.Dictionary and
'  ADODB.Stream are all late-bound (CreateObject), because a missing reference
'  breaks the ENTIRE VBA project, not just the module that needed it.
' ============================================================================

' ----------------------------------------------------------------------------
'  MODULE-LEVEL DECLARATIONS
'
'  These have to live up here, above the first procedure. VBA allows only
'  comments after an End Sub / End Function, so a Const parked next to the
'  procedure that uses it is a compile error, not a style choice.
' ----------------------------------------------------------------------------

' Excel caps a defined name's Comment at 255 characters and RAISES past it.
Private Const MAX_NAME_COMMENT As Long = 255

' Excel's hard ceiling on the length of a single formula. Nothing can be done
' about a definition longer than this -- but saying so beats reporting whatever
' generic error Excel happens to raise.
Private Const MAX_FORMULA_LEN As Long = 8192

' How much of Excel's error text to keep. Its descriptions run to several
' paragraphs of tutorial ("Not trying to type a formula? When the first
' character is an equal sign..."), which buries the actual reason and makes the
' report unreadable once more than two functions fail.
Private Const MAX_REASON_LEN As Long = 90


' ============================================================================
'  SOURCES -- each returns a Collection of Array(name, formula, description)
' ============================================================================

' Download and parse a GitHub Gist in Excel Labs / Advanced Formula Environment
' format. Accepts either the gist page URL or a raw URL.
Public Function EntriesFromGistUrl(ByVal url As String) As Collection
    Set EntriesFromGistUrl = ParseGistEntries(HttpGetText(NormalizeGistUrl(url)))
End Function

' Parse a local text file in the same format -- i.e. read back what
' ExportLibraryToTextFile wrote.
Public Function EntriesFromTextFile(ByVal path As String) As Collection
    Set EntriesFromTextFile = ParseGistEntries(ReadTextFile(path))
End Function

' Read a workbook exported by ExportLibraryToXlsx (or any workbook with
' Function / Formula / Description columns).
'
' Columns are matched BY HEADER NAME, never by position, so a file with the
' columns reordered still imports correctly. Every sheet is searched, and a real
' Excel table on the sheet is preferred over a bare header row.
Public Function EntriesFromXlsxFile(ByVal path As String) As Collection
    Dim out As New Collection
    Set EntriesFromXlsxFile = out

    Dim wb As Workbook
    On Error GoTo done

    ' ReadOnly + UpdateLinks:=0 so opening someone's backup can never prompt,
    ' recalculate, or modify the file being read.
    Set wb = Application.Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=True)

    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If ReadEntriesFromHeaderRow(lo.HeaderRowRange, out) Then GoTo done
        Next lo
    Next ws

    ' No table matched -- fall back to treating row 1 of each sheet as headers.
    For Each ws In wb.Worksheets
        If ws.UsedRange.rows.Count > 1 Then
            If ReadEntriesFromHeaderRow(ws.rows(1), out) Then GoTo done
        End If
    Next ws

done:
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number
    savedDesc = Err.description

    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If

    If savedErr <> 0 Then Err.Raise savedErr, "EntriesFromXlsxFile", savedDesc
End Function

' Find Function / Formula / Description in headerRow and append every data row
' beneath it to out. Returns True if the three headers were found.
Private Function ReadEntriesFromHeaderRow(ByVal headerRow As Range, _
                                          ByVal out As Collection) As Boolean
    Dim ixName As Long, ixCode As Long, ixDesc As Long
    Dim c As Range, h As String

    For Each c In headerRow.Cells
        h = LCase$(TrimWS(CStr(c.value & "")))
        Select Case h
            Case "function", "name":        If ixName = 0 Then ixName = c.Column
            Case "formula", "code":         If ixCode = 0 Then ixCode = c.Column
            Case "description", "comment":  If ixDesc = 0 Then ixDesc = c.Column
        End Select
    Next c

    If ixName = 0 Or ixCode = 0 Then Exit Function
    ReadEntriesFromHeaderRow = True

    Dim ws As Worksheet
    Set ws = headerRow.Worksheet

    Dim firstRow As Long, lastRow As Long, r As Long
    firstRow = headerRow.Row + 1
    lastRow = ws.Cells(ws.rows.Count, ixName).End(xlUp).Row
    If lastRow < firstRow Then Exit Function

    Dim nm As String, fx As String, ds As String
    For r = firstRow To lastRow
        nm = TrimWS(CStr(ws.Cells(r, ixName).value & ""))
        fx = TrimWS(CStr(ws.Cells(r, ixCode).value & ""))
        If ixDesc > 0 Then ds = CStr(ws.Cells(r, ixDesc).value & "") Else ds = ""

        If Len(nm) > 0 And Len(fx) > 0 Then
            ' Tolerate a formula stored without its leading "=".
            If Left$(fx, 1) <> "=" Then fx = "=" & fx
            out.Add Array(nm, fx, ds)
        End If
    Next r
End Function

' Every LAMBDA defined in a workbook's Name Manager, or just one of them.
'
' A LAMBDA is a defined name whose RefersTo starts with "=LAMBDA(" -- that test
' is the whole detection mechanism. Sheet-scoped names are skipped: their .Name
' reads "Sheet1!fxThing", which is not a usable function name, and a LAMBDA that
' only works on one sheet is not worth putting in a shared library.
Public Function EntriesFromWorkbookNames(ByVal wb As Workbook, _
                                         Optional ByVal onlyName As String = "") As Collection
    Dim out As New Collection
    Set EntriesFromWorkbookNames = out
    If wb Is Nothing Then Exit Function

    Dim nm As name, refers As String, cmt As String
    For Each nm In wb.names
        refers = ""
        cmt = ""

        ' A broken name raises on .RefersTo. Skip it rather than abandoning the
        ' whole scan -- one bad name should not cost the user the other fifty.
        On Error Resume Next
        refers = nm.RefersTo
        cmt = nm.comment
        On Error GoTo 0

        ' vbTextCompare, to agree with the test AddEntriesToWorkbook already uses.
        ' A plain = here is case-sensitive (no Option Compare Text in this module),
        ' so a name stored as "=lambda(" was invisible to the dropdown.
        If StrComp(Left$(refers, 8), "=LAMBDA(", vbTextCompare) = 0 Then
            If InStr(nm.name, "!") = 0 Then
                If Len(onlyName) = 0 Then
                    out.Add Array(nm.name, refers, cmt)
                ElseIf StrComp(nm.name, onlyName, vbTextCompare) = 0 Then
                    out.Add Array(nm.name, refers, cmt)
                    Exit For
                End If
            End If
        End If
    Next nm
End Function

' The stored library, or just one function from it.
Public Function EntriesFromLibrary(Optional ByVal onlyName As String = "") As Collection
    If Len(onlyName) = 0 Then
        Set EntriesFromLibrary = AddInStorage.LambdaAll()
    Else
        Dim out As New Collection
        Dim e As Variant
        e = AddInStorage.LambdaEntry(onlyName)
        If Not IsEmpty(e) Then out.Add e
        Set EntriesFromLibrary = out
    End If
End Function


' ============================================================================
'  SINKS
' ============================================================================

' Write entries into tblAlonzoChurch. Existing names are overwritten.
' Returns True on success; on False, AddInStorage.LastError says why.
'
' FORMULAS ARE FLATTENED ON THE WAY IN, the same as they are when assigned to a
' defined name. The two paths have to agree: if the library kept the author's
' indentation, a function could sit in the library looking fine and then be
' refused on inject for being over Excel's limit -- and what you saw stored
' would not be what Excel ends up holding. It also keeps the .xlam a third
' smaller, since the add-in saves this table to disk on every write.
Public Function AddEntriesToLibrary(ByVal entries As Collection, _
                                    ByRef addedCount As Long, _
                                    ByRef updatedCount As Long) As Boolean
    Dim flat As New Collection, e As Variant
    For Each e In entries
        flat.Add Array(e(0), FlattenFormula(CStr(e(1))), e(2))
    Next e

    AddEntriesToLibrary = AddInStorage.UpsertLambdas(flat, addedCount, updatedCount)
    If AddEntriesToLibrary Then AddInStorage.SaveStorage
End Function

' Write entries into a workbook's Name Manager. Returns the number added.
' Names that could not be created are appended to failed (as name strings).
'
' Formulas are flattened to one line on the way in (see FlattenFormula). The
' LIBRARY keeps the author's original layout -- only the copy handed to Excel is
' flattened -- so a text export still round-trips a readable, formatted Gist.
'
' REPEAT UNTIL NO PROGRESS, which is not belt-and-braces. A LAMBDA that calls
' another LAMBDA only validates once the name it calls already exists, so an
' import can fail purely on the order the definitions happen to appear in.
'
' A fixed two passes resolves a chain two deep and no further: with A calling B
' calling C, and C defined last, pass 1 adds only C and pass 2 only B. Looping
' while anything is still succeeding costs one extra sweep at the end and
' handles any depth. The counter is a runaway guard, nothing more -- the loop
' exits on its own the first time a sweep adds nothing.
'
' (No "= Nothing" on the optional parameter -- VBA only allows constant
' expressions as defaults, and an omitted object parameter is Nothing anyway.)
Public Function AddEntriesToWorkbook(ByVal entries As Collection, _
                                     ByVal wb As Workbook, _
                                     Optional ByVal failed As Collection, _
                                     Optional ByVal wrapped As Collection) As Long
    If wb Is Nothing Then Exit Function

    Dim done As Object, reasons As Object
    Set done = CreateObject("Scripting.Dictionary")
    Set reasons = CreateObject("Scripting.Dictionary")
    done.CompareMode = 1                          ' 1 = vbTextCompare
    reasons.CompareMode = 1

    Dim addedThisPass As Long, guard As Long, e As Variant, why As String
    Dim wasWrapped As Boolean
    Do
        addedThisPass = 0
        guard = guard + 1

        For Each e In entries
            If Not done.Exists(CStr(e(0))) Then
                why = ""
                wasWrapped = False
                If CreateLambdaName(wb, CStr(e(0)), CStr(e(1)), CStr(e(2)), why, wasWrapped) Then
                    done(CStr(e(0))) = True
                    addedThisPass = addedThisPass + 1
                    If wasWrapped And Not wrapped Is Nothing Then wrapped.Add CStr(e(0))
                Else
                    reasons(CStr(e(0))) = why      ' keep only the latest attempt
                End If
            End If
        Next e
    Loop While addedThisPass > 0 And done.Count < entries.Count And guard < 20

    ' Report WHY, not just WHICH. "FILL (error 1004 - ...)" points at the cause;
    ' a bare list of names sends you back to guessing.
    If Not failed Is Nothing Then
        For Each e In entries
            If Not done.Exists(CStr(e(0))) Then
                If Len(reasons(CStr(e(0)))) > 0 Then
                    failed.Add CStr(e(0)) & " (" & reasons(CStr(e(0))) & ")"
                Else
                    failed.Add CStr(e(0))
                End If
            End If
        Next e
    End If

    AddEntriesToWorkbook = done.Count
End Function

' Create or replace one defined name holding a LAMBDA. Returns True on success;
' on False, failReason carries Excel's own message so the caller can show it.
'
' NOTE: There is a bug in VBA.
' To ensure VBA works correctly in other regions with different separators,
' first need to clear/create the name with an empty string, add the comment,
' then add the function. Assigning RefersTo before the comment fails on any
' machine whose list separator is not a comma.
' WHY THIS DOES NOT JUST CREATE THE NAME AS "=0" AND THEN FILL IT IN
'
' Excel is perfectly happy with a formula that references a name which DOES NOT
' EXIST -- it resolves later, or evaluates to #NAME?. That is why a library can
' be imported in any order at all.
'
' What Excel will NOT accept is "FOO(x)" when FOO exists as something that is
' not a function. A placeholder of "=0" makes the name exist AS THE NUMBER ZERO,
' so every call to it becomes "0(x)" -- invalid. A recursive function assigned
' on top of its own "=0" placeholder is rejected by its own placeholder.
'
' Worse, the old code left that placeholder behind on failure, so one refusal
' poisoned every function that called it. That is the cascade seen in the edugca
' library: SPLIT was refused, and M, MULTIFILTER and TEXTTOARRAY -- the three
' functions that call SPLIT -- were then refused too, against a "SPLIT" that was
' sitting in the workbook as the number 0.
'
' So: create the name ALREADY CARRYING its formula, and if anything goes wrong,
' remove it. An absent name is harmless; a half-made one is contagious.
Private Function CreateLambdaName(ByVal wb As Workbook, ByVal functionName As String, _
                                  ByVal functionFormula As String, _
                                  ByVal functionComment As String, _
                                  Optional ByRef failReason As String, _
                                  Optional ByRef wrappedAsFunction As Boolean) As Boolean
    ' Flatten FIRST, then measure. The author's indentation is real characters
    ' as far as Excel's 8,192 limit is concerned, and it can be 40% of the text.
    ' Measuring before flattening would reject formulas that fit perfectly well.
    functionFormula = FlattenFormula(functionFormula)

    ' Only now is the length meaningful. Excel's own complaint about an
    ' over-long formula is a generic 1004 that says nothing useful.
    If Len(functionFormula) > MAX_FORMULA_LEN Then
        failReason = "formula is " & Format$(Len(functionFormula), "#,##0") & _
                     " characters even after removing layout; Excel's limit is " & _
                     Format$(MAX_FORMULA_LEN, "#,##0")
        Exit Function
    End If

    Dim nm As name, oldRefersTo As String, hadOld As Boolean

    On Error Resume Next
    Set nm = wb.names(functionName)
    If Not nm Is Nothing Then
        oldRefersTo = nm.RefersTo
        hadOld = (Len(oldRefersTo) > 0)

        ' DELETE, don't blank.
        '
        ' The previous version did "nm.RefersTo = """"" to clear an existing
        ' name before rewriting it. An empty string is not a valid formula, so
        ' Excel raises 1004 on that line. The code this was ported from hid it
        ' under a blanket On Error Resume Next; with real error handling it
        ' became fatal, and any name that ALREADY EXISTED was refused. That is
        ' why "About" + lambda failed on the second Hatmaker gist -- both gists
        ' define it, and the second import hit this path.
        '
        ' Deleting also means the new formula is assigned while the name is
        ' ABSENT, which is exactly what a recursive definition needs.
        nm.Delete
    End If
    Set nm = Nothing
    Err.Clear
    On Error GoTo 0

    ' --- attempt 1: create the name already carrying its formula ---
    Dim errNum As Long, errText As String
    On Error Resume Next
    wb.names.Add name:=functionName, RefersTo:=functionFormula
    errNum = Err.Number
    errText = Err.description
    Err.Clear
    On Error GoTo 0

    If errNum = 0 Then
        On Error Resume Next
        SetNameComment wb.names(functionName), functionComment
        On Error GoTo 0
        CreateLambdaName = True
        Exit Function
    End If

    ' --- attempt 2: the documented locale dance ---
    '
    ' NOTE: There is a bug in VBA. To ensure VBA works correctly in other
    ' regions with different separators, first create the name with a simple
    ' placeholder, add the comment, then add the real formula. "=0" is used
    ' rather than "" because "" is not a formula and Excel rejects it.
    On Error Resume Next
    Err.Clear
    wb.names.Add name:=functionName, RefersTo:="=0"
    If Err.Number = 0 Then
        SetNameComment wb.names(functionName), functionComment
        Err.Clear
        wb.names(functionName).RefersTo = functionFormula
        If Err.Number = 0 Then
            On Error GoTo 0
            CreateLambdaName = True
            Exit Function
        End If
    End If
    Err.Clear

    ' --- attempt 3: stop Excel having to evaluate it ---
    '
    ' A definition that is not a LAMBDA -- Craig Hatmaker's "About" + lambda is
    ' "=TRIM(TEXTSPLIT(...))" -- has to be EVALUATED by Excel when the name is
    ' created, where a LAMBDA is only stored. Past a certain size Excel refuses,
    ' with a bare 1004 and no explanation.
    '
    ' Measured on that function: the name, the app state, the two-step route and
    ' the _xlfn. prefix were each ruled out one at a time. Arrays are not the
    ' problem either -- "=SEQUENCE(3)" and "=TEXTSPLIT(""a,b"","","")" both store
    ' fine. Wrapping the formula in LAMBDA() was the only thing that worked, and
    ' it worked instantly.
    '
    ' The cost is real and the caller reports it: the name becomes a
    ' ZERO-ARGUMENT FUNCTION, so it is called as Name() rather than Name. The
    ' result is identical. Better a function that needs two extra characters
    ' than no function at all.
    If InStr(1, functionFormula, "=LAMBDA(", vbTextCompare) <> 1 Then
        On Error Resume Next
        Err.Clear
        wb.names.Add name:=functionName, _
                     RefersTo:="=LAMBDA(" & Mid$(functionFormula, 2) & ")"
        If Err.Number = 0 Then
            SetNameComment wb.names(functionName), functionComment
            Err.Clear
            On Error GoTo 0
            wrappedAsFunction = True
            CreateLambdaName = True
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If

    ' --- all failed: leave the workbook as we found it ---
    '
    ' A name left stranded at "=0" is worse than no name at all: every OTHER
    ' function that calls it then fails too, because "FOO(x)" where FOO is the
    ' number zero is not a valid formula. One refusal would cascade.
    wb.names(functionName).Delete
    If hadOld Then wb.names.Add name:=functionName, RefersTo:=oldRefersTo
    Err.Clear
    On Error GoTo 0

    failReason = "error " & errNum & " - " & OneLine(errText, MAX_REASON_LEN)
End Function

' Collapse every run of whitespace OUTSIDE a string literal to a single space.
'
' THIS IS WHAT EXCEL LABS DOES, and it is not cosmetic tidying. A Gist formula
' carries the author's indentation, and Excel counts every one of those
' characters against its 8,192 limit. In Craig Hatmaker's debt library the
' indentation is 39% of the text:
'
'     FlexLoan   10,024 chars as written  ->  6,628 flattened
'     Revolver    9,450                   ->  5,296
'
' Both are refused at their written size and accepted at their flattened size.
' The formula is identical either way; only the layout differs.
'
' WHITESPACE INSIDE STRING LITERALS IS UNTOUCHABLE. Craig aligns his
' documentation table with runs of spaces and em-spaces inside the strings --
' collapsing those would wreck the output the function exists to produce. The
' quote tracking below is what keeps them intact.
'
' One space is kept rather than none, so two tokens can never fuse together.
Private Function FlattenFormula(ByVal s As String) As String
    Dim n As Long: n = Len(s)
    If n = 0 Then Exit Function

    Dim out As String: out = Space$(n)
    Dim i As Long, p As Long, inQuote As Boolean, ch As String

    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If inQuote Then
            p = p + 1: Mid$(out, p, 1) = ch
            If ch = """" Then inQuote = False
            i = i + 1
        ElseIf ch = """" Then
            inQuote = True
            p = p + 1: Mid$(out, p, 1) = ch
            i = i + 1
        ElseIf ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            Do While i <= n
                ch = Mid$(s, i, 1)
                If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
                    i = i + 1
                Else
                    Exit Do
                End If
            Loop
            If p > 0 Then
                If Mid$(out, p, 1) <> " " Then
                    p = p + 1: Mid$(out, p, 1) = " "
                End If
            End If
        Else
            p = p + 1: Mid$(out, p, 1) = ch
            i = i + 1
        End If
    Loop

    FlattenFormula = TrimWS(Left$(out, p))
End Function

' Collapse a multi-line message to a single trimmed line and cap its length.
Private Function OneLine(ByVal text As String, ByVal maxLen As Long) As String
    Dim s As String
    s = Replace(Replace(Replace(text, vbCrLf, " "), vbLf, " "), vbCr, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    If Len(s) > maxLen Then s = Left$(s, maxLen) & "..."
    OneLine = s
End Function

' Set a name's comment, tolerating anything that goes wrong.
'
' Excel caps Name.Comment at 255 characters and RAISES if you exceed it. The
' description is decoration; losing it must never cost you the function. So it
' is truncated first and its failure is swallowed.
'
' This is why "Primes" disappeared from the Vertex42 import: its doc comment is
' 292 characters, the assignment raised, and the caller's error handler threw
' away a perfectly good 6,658-character formula on account of its caption.
'
' On Error Resume Next is scoped to this Sub, so the caller's own handler is
' back in force the moment it returns.
Private Sub SetNameComment(ByVal nm As name, ByVal text As String)
    On Error Resume Next
    If Len(text) > MAX_NAME_COMMENT Then text = Left$(text, MAX_NAME_COMMENT - 3) & "..."
    nm.comment = text
End Sub


' ============================================================================
'  EXPORTS
' ============================================================================

' Write the library to a new .xlsx at path: one sheet, one 3-column table, the
' same headers EntriesFromXlsxFile reads back.
'
' Takes the path rather than showing a Save dialog. The dialogs live in
' modLambdaRibbon, so everything in this module can be driven from the
' Immediate window -- which, with no test harness in the project, is the only
' way to exercise this code without clicking through the ribbon:
'
'     modLambdaLib.ExportLibraryToXlsx "C:\Temp\lib.xlsx"
'     ? modLambdaLib.EntriesFromXlsxFile("C:\Temp\lib.xlsx").Count
Public Sub ExportLibraryToXlsx(ByVal path As String)
    Dim entries As Collection
    Set entries = AddInStorage.LambdaAll()

    Dim wb As Workbook, ws As Worksheet
    On Error GoTo CleanUp

    ' FastModeOn captures the user's real prior state and FastModeOff puts it
    ' back, so someone working in manual calculation is not silently switched to
    ' automatic. DisplayAlerts is set to False on top of it because
    ' GetSaveAsFilename has already asked about overwriting -- without this,
    ' SaveAs asks a second time.
    AppStateManager.FastModeOn
    Application.DisplayAlerts = False

    ' xlWBATWorksheet = a one-sheet workbook regardless of the user's
    ' "sheets in new workbook" setting.
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.name = "LAMBDA library"

    ws.Range("A1").value = "Function"
    ws.Range("B1").value = "Formula"
    ws.Range("C1").value = "Description"

    Dim n As Long
    n = entries.Count

    If n > 0 Then
        Dim buf() As Variant
        ReDim buf(1 To n, 1 To 3)
        Dim i As Long, e As Variant
        For i = 1 To n
            e = entries(i)
            buf(i, 1) = e(0)
            buf(i, 2) = e(1)
            buf(i, 3) = e(2)
        Next i

        ' Text format BEFORE the values, or Excel tries to evaluate "=LAMBDA(..."
        ' as a formula on the way in and the export is full of #NAME? errors.
        ws.Range("B2").Resize(n, 1).NumberFormat = "@"
        ws.Range("A2").Resize(n, 3).value = buf
    End If

    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(IIf(n = 0, 2, n + 1), 3), , xlYes)
    lo.name = "tblAlonzoChurch"

    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 70
    ws.Columns("C").ColumnWidth = 50
    ws.Range("B:C").WrapText = True
    ws.rows(1).Font.bold = True

    ' 51 = xlOpenXMLWorkbook (.xlsx)
    wb.SaveAs fileName:=path, FileFormat:=51
    wb.Close SaveChanges:=False

CleanUp:
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number
    savedDesc = Err.description

    On Error Resume Next
    If savedErr <> 0 And Not wb Is Nothing Then wb.Close SaveChanges:=False
    AppStateManager.FastModeOff
    On Error GoTo 0

    If savedErr <> 0 Then Err.Raise savedErr, "ExportLibraryToXlsx", savedDesc
End Sub

' Write the library to path as Gist-format text -- the same thing
' EntriesFromTextFile and EntriesFromGistUrl read.
Public Sub ExportLibraryToTextFile(ByVal path As String)
    WriteTextFile path, FormatAsGistText(AddInStorage.LambdaAll())
End Sub

' The exact inverse of ParseGistEntries: a doc comment carrying the
' description, then "NAME = formula;".
'
' The leading "=" is stripped here because the format expects
' "NAME = LAMBDA(...)", not "NAME = =LAMBDA(...)". ProcessEntry puts it back on
' the way in. Getting this pair wrong is the one thing that would stop the
' export round-tripping, which is why the export/import round trip is the check
' worth running after any change to either function.
Private Function FormatAsGistText(ByVal entries As Collection) As String
    Dim sb As String, e As Variant
    Dim fx As String, ds As String, lines() As String, i As Long

    For Each e In entries
        ds = CStr(e(2))
        If Len(TrimWS(ds)) > 0 Then
            sb = sb & "/**" & vbCrLf
            ds = Replace(Replace(ds, vbCrLf, vbLf), vbCr, vbLf)
            lines = Split(ds, vbLf)
            For i = LBound(lines) To UBound(lines)
                sb = sb & " * " & lines(i) & vbCrLf
            Next i
            sb = sb & " */" & vbCrLf
        End If

        fx = CStr(e(1))
        If Left$(fx, 1) = "=" Then fx = Mid$(fx, 2)

        sb = sb & CStr(e(0)) & " = " & fx & ";" & vbCrLf & vbCrLf
    Next e

    FormatAsGistText = sb
End Function


' ============================================================================
'  GIST PARSING
' ============================================================================

' Turn a gist page URL into a raw URL. Passes raw URLs through unchanged.
Public Function NormalizeGistUrl(ByVal url As String) As String
    Dim u As String: u = url
    Dim p As Long
    p = InStr(u, "#"): If p > 0 Then u = Left$(u, p - 1)
    p = InStr(u, "?"): If p > 0 Then u = Left$(u, p - 1)
    Do While Right$(u, 1) = "/": u = Left$(u, Len(u) - 1): Loop

    If InStr(1, u, "gist.github.com", vbTextCompare) > 0 And _
       InStr(1, u, "githubusercontent", vbTextCompare) = 0 Then
        u = Replace(u, "gist.github.com", "gist.githubusercontent.com", 1, 1, vbTextCompare)
        If InStr(1, u, "/raw", vbTextCompare) = 0 Then u = u & "/raw"
    End If
    NormalizeGistUrl = u
End Function

Private Function HttpGetText(ByVal url As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", "Excel-LAMBDA-Importer"
    http.Send
    If http.Status <> 200 Then
        Err.Raise vbObjectError + 513, "HttpGetText", _
            "Download failed: HTTP " & http.Status & " " & http.StatusText & vbCrLf & url
    End If
    HttpGetText = http.responseText
End Function

' ----------------------------------------------------------------------------
'  Parse the source into Array(name, "=formula", description).
'
'  The scan is QUOTE- AND PAREN-AWARE, which is the only reason it works: a
'  LAMBDA body is full of commas, parens and quoted strings, and a naive
'  Split on ";" would cut a formula in half the moment one appeared inside a
'  text literal. Definitions are separated by a ";" seen at paren depth 0 and
'  outside quotes; name and formula are split on the first "=" under the same
'  rule. Line breaks inside a formula are preserved.
'
'  A "/** ... */" block immediately before a definition becomes its description.
'  Plain "/* ... */" blocks (section banners) are skipped.
' ----------------------------------------------------------------------------
Public Function ParseGistEntries(ByVal src As String) As Collection
    Dim coll As New Collection
    Dim n As Long: n = Len(src)
    Dim i As Long, depth As Long, inQuote As Boolean, chunkStart As Long
    Dim pendingComment As String, ch As String, j As Long, block As String
    chunkStart = 1: i = 1
    Do While i <= n
        ch = Mid$(src, i, 1)
        If inQuote Then
            If ch = """" Then inQuote = False
            i = i + 1
        ElseIf ch = """" Then
            inQuote = True: i = i + 1
        ElseIf ch = "/" And i < n And Mid$(src, i + 1, 1) = "/" Then
            ' Skip a // line comment HERE, not just when building the formula.
            ' A ";" or an unbalanced "(" inside a comment would otherwise be
            ' counted and split a definition in the wrong place.
            Do While i <= n
                ch = Mid$(src, i, 1)
                If ch = vbCr Or ch = vbLf Then Exit Do
                i = i + 1
            Loop
        ElseIf depth = 0 And ch = "/" And i < n And Mid$(src, i + 1, 1) = "*" Then
            j = InStr(i + 2, src, "*/")
            If j = 0 Then
                block = Mid$(src, i): i = n + 1
            Else
                block = Mid$(src, i, j + 2 - i): i = j + 2
            End If
            If Left$(block, 3) = "/**" Then pendingComment = CleanDocComment(block)
            chunkStart = i                       ' comment is not part of the next definition
        ElseIf ch = "/" And i < n And Mid$(src, i + 1, 1) = "*" Then
            ' A block comment INSIDE a definition body. Skip it, but leave
            ' chunkStart alone -- the definition started before this.
            j = InStr(i + 2, src, "*/")
            If j = 0 Then i = n + 1 Else i = j + 2
        ElseIf ch = "(" Then
            depth = depth + 1: i = i + 1
        ElseIf ch = ")" Then
            If depth > 0 Then depth = depth - 1
            i = i + 1
        ElseIf ch = ";" And depth = 0 Then
            ProcessEntry Mid$(src, chunkStart, i - chunkStart), pendingComment, coll
            pendingComment = ""
            chunkStart = i + 1: i = i + 1
        Else
            i = i + 1
        End If
    Loop
    If chunkStart <= n Then ProcessEntry Mid$(src, chunkStart, n - chunkStart + 1), pendingComment, coll
    Set ParseGistEntries = coll
End Function

Private Sub ProcessEntry(ByVal chunk As String, ByVal description As String, _
                         ByVal coll As Collection)
    ' Comments are legal in the Gist format and ILLEGAL in an Excel formula.
    ' They have to come out before the text is ever handed to Names.Add.
    Dim t As String: t = TrimWS(StripComments(chunk))
    If Len(t) = 0 Then Exit Sub

    Dim i As Long, depth As Long, inQuote As Boolean, eqPos As Long, ch As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If inQuote Then
            If ch = """" Then inQuote = False
        Else
            Select Case ch
                Case """": inQuote = True
                Case "(": depth = depth + 1
                Case ")": If depth > 0 Then depth = depth - 1
                Case "="
                    If depth = 0 Then eqPos = i: Exit For
            End Select
        End If
    Next i
    If eqPos = 0 Then Exit Sub

    Dim nm As String, code As String
    nm = TrimWS(Left$(t, eqPos - 1))
    code = TrimWS(Mid$(t, eqPos + 1))
    If Len(nm) = 0 Or Len(code) = 0 Then Exit Sub
    If Not IsValidName(nm) Then Exit Sub

    code = Replace(code, vbCrLf, vbLf)           ' in-cell line breaks
    code = Replace(code, vbCr, vbLf)
    coll.Add Array(nm, "=" & code, description)
End Sub

' Remove // line comments and /* */ block comments from a definition, leaving
' quoted strings untouched.
'
' THIS IS WHY 12 OF THE 118 VERTEX42 FUNCTIONS USED TO FAIL. The Advanced
' Formula Environment format allows JavaScript-style comments inside a formula;
' an Excel formula does not. Leave a "// Handle defaults" in the text and
' Names.Add rejects the whole definition. Excel Labs strips them on import,
' which is exactly why it succeeded where this add-in did not.
'
' The quote awareness is not a nicety. Nearly every one of those functions
' embeds its own documentation URL:
'
'     LET(doc, "https://www.vertex42.com/lambda/reparray.html",
'
' A naive search for "//" would cut the formula in half at the "https://".
'
' Line breaks are kept -- only the comment text goes -- so the formula still
' reads sensibly in the formula bar and in the library's Formula column.
Private Function StripComments(ByVal s As String) As String
    Dim n As Long: n = Len(s)
    If n = 0 Then Exit Function

    ' Preallocate and fill with Mid$. Building the result with "out = out & ch"
    ' reallocates the whole string on every character, which on a 100KB Gist is
    ' the difference between instant and a visible freeze.
    Dim out As String: out = Space$(n)
    Dim i As Long, p As Long, j As Long, inQuote As Boolean, ch As String

    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If inQuote Then
            p = p + 1: Mid$(out, p, 1) = ch
            If ch = """" Then inQuote = False
            i = i + 1
        ElseIf ch = """" Then
            inQuote = True
            p = p + 1: Mid$(out, p, 1) = ch
            i = i + 1
        ElseIf ch = "/" And i < n And Mid$(s, i + 1, 1) = "/" Then
            Do While i <= n
                ch = Mid$(s, i, 1)
                If ch = vbCr Or ch = vbLf Then Exit Do   ' keep the line break
                i = i + 1
            Loop
        ElseIf ch = "/" And i < n And Mid$(s, i + 1, 1) = "*" Then
            j = InStr(i + 2, s, "*/")
            If j = 0 Then i = n + 1 Else i = j + 2
        Else
            p = p + 1: Mid$(out, p, 1) = ch
            i = i + 1
        End If
    Loop

    StripComments = Left$(out, p)
End Function

' Clean a /** ... */ doc comment into text (lines joined with line breaks).
Private Function CleanDocComment(ByVal block As String) As String
    Dim s As String: s = block
    If Left$(s, 3) = "/**" Then
        s = Mid$(s, 4)
    ElseIf Left$(s, 2) = "/*" Then
        s = Mid$(s, 3)
    End If
    If Right$(s, 2) = "*/" Then s = Left$(s, Len(s) - 2)
    s = Replace(s, vbCrLf, vbLf): s = Replace(s, vbCr, vbLf)

    Dim parts() As String: parts = Split(s, vbLf)
    Dim out As String, k As Long, ln As String
    For k = LBound(parts) To UBound(parts)
        ln = TrimWS(parts(k))
        Do While Left$(ln, 1) = "*"
            ln = TrimWS(Mid$(ln, 2))
        Loop
        If Len(ln) > 0 Then
            If Len(out) > 0 Then out = out & vbLf
            out = out & ln
        End If
    Next k
    CleanDocComment = out
End Function


' ============================================================================
'  MODULE NAMESPACING (Excel Labs style)
' ============================================================================

' Prefix every function with "module." and rewrite references BETWEEN the
' imported functions to match -- so importing under "VLL" turns RESCALE into
' VLL.RESCALE everywhere, including inside the formulas that call it.
Public Function ApplyModuleNamespace(ByVal entries As Collection, _
                                     ByVal moduleName As String) As Collection
    Dim known As Object
    Set known = CreateObject("Scripting.Dictionary")
    known.CompareMode = vbBinaryCompare           ' exact-spelling match

    Dim e As Variant
    For Each e In entries: known(CStr(e(0))) = True: Next e

    Dim out As New Collection
    For Each e In entries
        out.Add Array(moduleName & "." & CStr(e(0)), _
                      PrefixReferences(CStr(e(1)), known, moduleName), _
                      e(2))
    Next e
    Set ApplyModuleNamespace = out
End Function

' Rewrite whole-token references to known names. Quote-aware, so a function
' name that also appears inside a text literal is left alone.
Private Function PrefixReferences(ByVal formula As String, ByVal known As Object, _
                                  ByVal moduleName As String) As String
    Dim n As Long: n = Len(formula)
    Dim res As String, lastCopied As Long: lastCopied = 1
    Dim i As Long, inQuote As Boolean, ch As String, startPos As Long, tok As String
    i = 1
    Do While i <= n
        ch = Mid$(formula, i, 1)
        If inQuote Then
            If ch = """" Then inQuote = False
            i = i + 1
        ElseIf ch = """" Then
            inQuote = True
            i = i + 1
        ElseIf IsNameStart(ch) Then
            startPos = i
            Do While i <= n
                If IsNameChar(Mid$(formula, i, 1)) Then i = i + 1 Else Exit Do
            Loop
            tok = Mid$(formula, startPos, i - startPos)
            If known.Exists(tok) Then
                res = res & Mid$(formula, lastCopied, startPos - lastCopied) & moduleName & "." & tok
                lastCopied = i
            End If
        Else
            i = i + 1
        End If
    Loop
    res = res & Mid$(formula, lastCopied)
    PrefixReferences = res
End Function


' ============================================================================
'  SMALL SHARED HELPERS
' ============================================================================

' True for any character above U+007F.
'
' Excel defined names are UNICODE, and "ch Like [A-Za-z]" is not. Treating
' anything non-ASCII as a letter is deliberately loose -- VBA has no Unicode
' character-class test, and Excel does the real validation when the name is
' created, so the worst case is that a bad name is reported as skipped instead
' of rejected up front.
'
' AscW returns a SIGNED Integer, so code points above U+7FFF come back negative.
' Testing only "> 127" would silently drop those.
Private Function IsHighChar(ByVal ch As String) As Boolean
    Dim w As Integer
    w = AscW(ch)
    IsHighChar = (w > 127) Or (w < 0)
End Function

Private Function IsNameStart(ByVal ch As String) As Boolean
    IsNameStart = (ch Like "[A-Za-z]") Or (ch = "_") Or IsHighChar(ch)
End Function

Private Function IsNameChar(ByVal ch As String) As Boolean
    IsNameChar = (ch Like "[A-Za-z0-9_]") Or IsHighChar(ch)
End Function

' Is this usable as an Excel defined name?
'
' "." IS ALLOWED, deliberately: Excel Labs namespacing produces VLL.RESCALE, and
' a validator that rejects the dot would reject every namespaced import.
'
' SO ARE NON-ASCII LETTERS, and that matters more than it sounds. Craig
' Hatmaker's widely used "5G" libraries name every single function with a
' trailing GREEK SMALL LETTER LAMBDA, U+03BB -- "About" + lambda, "Timeline" +
' lambda, and so on. An ASCII-only test rejected all of them, so those Gists
' imported NOTHING, and the failure looked like "the parser is broken" rather
' than "the name test is too strict". Excel accepts these names happily; the
' validator was the only obstacle.
'
' (The character itself is deliberately NOT written in this comment: a .bas file
' is Windows-1252 and U+03BB has no representation in it, so spelling it out
' here would corrupt on the next export/import round trip.)
Public Function IsValidName(ByVal nm As String) As Boolean
    If Len(nm) = 0 Or Len(nm) > 255 Then Exit Function
    If Not IsNameStart(Left$(nm, 1)) Then Exit Function
    Dim i As Long, ch As String
    For i = 2 To Len(nm)
        ch = Mid$(nm, i, 1)
        If Not (IsNameChar(ch) Or ch = ".") Then Exit Function
    Next i
    IsValidName = True
End Function

' Full whitespace trim. VBA's Trim$ strips only spaces -- not CR, LF or Tab --
' so a name preceded by a newline survives Trim$ and then fails IsValidName.
Public Function TrimWS(ByVal s As String) As String
    Dim a As Long, b As Long, n As Long: n = Len(s)
    a = 1
    Do While a <= n
        Select Case Mid$(s, a, 1)
            Case " ", vbCr, vbLf, vbTab: a = a + 1
            Case Else: Exit Do
        End Select
    Loop
    b = n
    Do While b >= a
        Select Case Mid$(s, b, 1)
            Case " ", vbCr, vbLf, vbTab: b = b - 1
            Case Else: Exit Do
        End Select
    Loop
    If b < a Then TrimWS = "" Else TrimWS = Mid$(s, a, b - a + 1)
End Function

' Read a whole text file as UTF-8.
'
' ADODB.Stream rather than VBA's Open/Input because Open reads the machine's ANSI
' codepage: a file containing a curly quote or an en-dash comes back as mojibake
' on one machine and fine on another. Late-bound, so no project reference.
Private Function ReadTextFile(ByVal path As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                                  ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    ReadTextFile = stm.ReadText(-1)               ' adReadAll
    stm.Close
End Function

' Write a whole text file as UTF-8, without a byte-order mark.
'
' The BOM dance is not optional: ADODB.Stream always writes one for utf-8, and a
' BOM at the head of the file lands on the first function name, so re-importing
' the file you just exported silently loses its first definition.
Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    Dim stm As Object, bin As Object

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                                  ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText content

    ' Re-read as bytes, skip the 3 BOM bytes, save that.
    stm.Position = 0
    stm.Type = 1                                  ' adTypeBinary
    stm.Position = 3

    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1
    bin.Open
    stm.CopyTo bin
    bin.SaveToFile path, 2                        ' 2 = adSaveCreateOverWrite
    bin.Close
    stm.Close
End Sub
