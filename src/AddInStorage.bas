Attribute VB_Name = "AddInStorage"
Option Explicit

' ============================================================================
'  AddInStorage
'  The ONE module that knows where the XL Edge add-in keeps its OWN data.
'  Drop this standard module into the .XLAM add-in once.
'
'  Why this module exists:
'   * The add-in stores its own lists, settings, and LAMBDA functions on a
'     HIDDEN "reference" worksheet inside THISWORKBOOK (the add-in) -- NOT in
'     ActiveWorkbook (the user's file the macros operate on).
'   * Qualifying add-in storage to ActiveWorkbook is a top-tier add-in bug: it
'     errors (the user's file has no such sheet) or silently hits a same-named
'     sheet the user happens to have open. Everything here uses ThisWorkbook.
'   * Every storage reference is funneled through the accessors below, so a
'     sheet or table rename is a one-line change here, not a hunt across every
'     macro in the project.
'
'  Usage from any macro:
'     Dim names As Variant, i As Long
'     names = AddInStorage.CompanyList()        ' user-customizable, was hardcoded
'     For i = LBound(names) To UBound(names)
'         Debug.Print names(i)
'     Next i
'
'     Dim fx As Collection: Set fx = AddInStorage.LambdaAll()
'     AddInStorage.UpsertLambda "fxNet", "=LAMBDA(a,b,a-b)", "Difference."
'     AddInStorage.SaveStorage                  ' writers do NOT save for you
'
'  LAMBDA LIBRARY (tblAlonzoChurch)
'   * Three columns: Function | Formula | Description. See the LAMBDA LIBRARY
'     section near the bottom for the full read/write API and for why every
'     writer rewrites the whole table instead of adding rows.
'
'  SETTINGS (tblConstants)
'   * The values that used to be `Const` declarations in modWork_Code,
'     modFinancial_Standards and modProductivity_Code now live in tblConstants
'     on the reference sheet, so an end user can change them from the settings
'     form without opening the VBA editor.
'   * They are exposed below as Public Property Get procedures NAMED EXACTLY
'     LIKE THE OLD CONSTANTS. In VBA a Public Property Get in a standard module
'     has the same global visibility a Public Const had, so every existing call
'     site -- PERIOD_END, DEFAULT_ROW_HEIGHT, ... -- compiles and runs unchanged.
'     Migrating meant deleting the old declarations, nothing else.
'   * Every property falls back to a private DEF_* default (the value that used
'     to be hardcoded). A missing sheet, a missing key, a blank cell or a
'     wrong-typed cell degrades to today's behavior instead of breaking the
'     add-in.
' ============================================================================

' ----------------------------------------------------------------------------
'  EDIT ONLY THIS BLOCK to match the workbook.
' ----------------------------------------------------------------------------
Private Const REFERENCE_SHEET     As String = "shtReference"    ' hidden reference sheet
Private Const LAMBDA_TABLE        As String = "tblAlonzoChurch" ' table of LAMBDA definitions
Private Const COMPANY_TABLE       As String = "tblCompany"      ' list of company/department names
Private Const COMPANY_COLUMN      As Variant = "Company"        ' column header (or 1-based index) to read

Private Const CONSTANTS_TABLE     As String = "tblConstants"    ' user-editable settings
Private Const CONST_KEY_COLUMN    As Variant = "Constants"      ' column holding the setting NAME
Private Const CONST_VAL_COLUMN    As Variant = "Value"          ' column holding the setting VALUE

' The three columns of tblAlonzoChurch. A "Tags" column used to sit alongside
' these; it was dropped because nobody filled it in. Everything is a LAMBDA --
' there is deliberately no "Type" column.
Private Const LAMBDA_NAME_COLUMN  As Variant = "Function"       ' the function name
Private Const LAMBDA_CODE_COLUMN  As Variant = "Formula"        ' "=LAMBDA(...)" stored as TEXT
Private Const LAMBDA_DESC_COLUMN  As Variant = "Description"    ' what the function does

' ----------------------------------------------------------------------------
'  FALLBACK DEFAULTS -- the values that were hardcoded before tblConstants.
'  These are the safety net, NOT the source of truth. Change a value for real
'  in the settings form; change it here only to move the "if all else fails"
'  floor. Named DEF_* so they don't collide with the Property Get names below.
' ----------------------------------------------------------------------------
Private Const DEF_PERIOD_END               As Date = #6/30/2026#
Private Const DEF_ITD_START                As Date = #1/1/2019#
Private Const DEF_FORMULABAR_HEIGHT_SM     As Double = 2
Private Const DEF_FORMULABAR_HEIGHT_LG     As Double = 15
Private Const DEF_LOWER_BOUND              As Double = 1
Private Const DEF_UPPER_BOUND              As Double = 100
Private Const DEF_HEADER_ROW_HEIGHT_SINGLE As Double = 42
Private Const DEF_HEADER_ROW_HEIGHT_MULTI  As Double = 32
Private Const DEF_DEFAULT_ROW_HEIGHT       As Double = 15
Private Const DEF_DEFAULT_COLUMN_WIDTH     As Double = 15
Private Const DEF_FOOTER_FONT_SIZE         As Double = 8
Private Const DEF_MYFOOTER                 As String = _
    "This report contains confidential information. " & _
    "Unauthorized disclosure or distribution is prohibited."

' ----------------------------------------------------------------------------
'  Module-level cache. Reading a worksheet cell is a COM call; the settings are
'  read many times per run and change rarely, so they are loaded once into a
'  Dictionary and reused.
'
'  Late-bound (CreateObject) on purpose -- early binding would require every
'  machine running this add-in to carry a reference to Microsoft Scripting
'  Runtime, and a missing reference breaks the whole VBA project, not just
'  this module.
'
'  Anything that writes to the reference sheet MUST call InvalidateCache.
' ----------------------------------------------------------------------------
Private mCache As Object            ' Scripting.Dictionary; Nothing = not loaded

' Last failure from a writer, for the settings form to display.
'
' The writers return True/False rather than raising, so the form can decide how
' to report. Without stashing the text here that detail was simply lost -- a
' "could not be written" message with no error number is close to useless.
Private mLastError As String

' Human-readable reason the last SetConstant / SetCompanyList returned False.
Public Property Get LastError() As String
    LastError = mLastError
End Property

' ============================================================================
'  Sheet / table accessors -- ALWAYS from ThisWorkbook, never ActiveWorkbook.
' ============================================================================

' The add-in's hidden reference sheet (lists + LAMBDA table live here).
Public Function ReferenceSheet() As Worksheet
    Set ReferenceSheet = ThisWorkbook.Worksheets(REFERENCE_SHEET)
End Function

' The table of LAMBDA definitions.
Public Function LambdaTable() As ListObject
    Set LambdaTable = ReferenceSheet().ListObjects(LAMBDA_TABLE)
End Function

' The user-customizable list of company / department names.
Public Function CompanyTable() As ListObject
    Set CompanyTable = ReferenceSheet().ListObjects(COMPANY_TABLE)
End Function

' The user-customizable settings table.
Public Function ConstantsTable() As ListObject
    Set ConstantsTable = ReferenceSheet().ListObjects(CONSTANTS_TABLE)
End Function

' ============================================================================
'  List readers -- turn a table column into a VBA array to loop over.
' ============================================================================

' Read one column of an Excel table into a 1-D, 1-based array of its values.
' colRef is the column header (String) or its 1-based position (Long).
' Blank trailing cells are dropped. Returns a 0-length array if the table is
' empty -- callers should guard: If HasItems(arr) Then ...
Public Function TableColumnList(ByVal lo As ListObject, ByVal colRef As Variant) As Variant
    If lo.DataBodyRange Is Nothing Then
        TableColumnList = Array()                  ' no rows -> empty array
        Exit Function
    End If

    Dim lc As ListColumn
    If VarType(colRef) = vbString Then
        Set lc = lo.ListColumns(CStr(colRef))
    Else
        Set lc = lo.ListColumns(CLng(colRef))
    End If

    Dim data As Variant
    data = lc.DataBodyRange.Value2                  ' (rows, 1) 2-D, or scalar if 1 row

    Dim out() As Variant, n As Long, i As Long, k As Long
    If Not IsArray(data) Then                       ' single-row table -> scalar
        If Len(CStr(data)) = 0 Then
            TableColumnList = Array()
        Else
            ReDim out(1 To 1): out(1) = data
            TableColumnList = out
        End If
        Exit Function
    End If

    n = UBound(data, 1)
    ReDim out(1 To n)
    k = 0
    For i = 1 To n
        If Len(CStr(data(i, 1))) > 0 Then           ' skip blanks
            k = k + 1
            out(k) = data(i, 1)
        End If
    Next i

    If k = 0 Then
        TableColumnList = Array()
    Else
        ReDim Preserve out(1 To k)
        TableColumnList = out
    End If
End Function

' The company/department names, ready to loop. Replaces the hardcoded list.
Public Function CompanyList() As Variant
    CompanyList = TableColumnList(CompanyTable(), COMPANY_COLUMN)
End Function

' True if an array returned by the list readers actually has at least one item.
Public Function HasItems(ByVal arr As Variant) As Boolean
    If IsArray(arr) Then HasItems = (UBound(arr) >= LBound(arr))
End Function

' True if the reference sheet can be resolved -- call this in a guard at the
' top of macros that depend on add-in storage.
Public Function StorageReady() As Boolean
    On Error GoTo NotReady
    Dim ws As Worksheet
    Set ws = ReferenceSheet()
    StorageReady = Not ws Is Nothing
    Exit Function
NotReady:
    StorageReady = False
End Function

' ============================================================================
'  Settings cache
' ============================================================================

' Drop the cached settings so the next read comes off the sheet again.
' Call after ANY write to tblConstants.
Public Sub InvalidateCache()
    Set mCache = Nothing
End Sub

' Load tblConstants into the cache. Silent on failure: mCache is left as an
' empty Dictionary so every lookup misses and every caller gets its DEF_*
' fallback. That is the whole point -- a damaged reference sheet must not stop
' the add-in from running.
Private Sub LoadCache()
    Set mCache = CreateObject("Scripting.Dictionary")
    mCache.CompareMode = 1                          ' 1 = vbTextCompare, case-insensitive keys

    On Error GoTo done                              ' any failure -> empty cache -> defaults

    Dim lo As ListObject
    Set lo = ConstantsTable()
    If lo.DataBodyRange Is Nothing Then GoTo done   ' table exists but has no rows

    ' ONE read of each column instead of a COM call per cell.
    '
    ' .Value, NOT .Value2, on the value column: .Value2 returns a date as its
    ' raw serial number (6/30/2026 -> 46203). That would not error -- it would
    ' quietly hand back a number where a Date was expected, which is the kind
    ' of bug that shows up as a wrong report rather than a crash.
    Dim keys As Variant, vals As Variant
    keys = lo.ListColumns(CONST_KEY_COLUMN).DataBodyRange.value
    vals = lo.ListColumns(CONST_VAL_COLUMN).DataBodyRange.value

    Dim i As Long, n As Long, k As String
    If Not IsArray(keys) Then                       ' single-row table -> scalars
        k = Trim$(CStr(keys))
        If Len(k) > 0 Then mCache(k) = vals
        GoTo done
    End If

    n = UBound(keys, 1)
    For i = 1 To n
        k = Trim$(CStr(keys(i, 1)))
        If Len(k) > 0 Then mCache(k) = vals(i, 1)   ' last duplicate key wins
    Next i

done:
    On Error GoTo 0
End Sub

' Raw cached value for a key, or Empty if absent/blank. Callers should prefer
' the typed readers below.
Public Function ConstantValue(ByVal key As String) As Variant
    If mCache Is Nothing Then LoadCache
    If mCache.Exists(key) Then
        If Len(Trim$(CStr(mCache(key)))) > 0 Then
            ConstantValue = mCache(key)
            Exit Function
        End If
    End If
    ConstantValue = Empty
End Function

' True if the key exists in the table with a non-blank value.
Public Function HasConstant(ByVal key As String) As Boolean
    HasConstant = Not IsEmpty(ConstantValue(key))
End Function

' Every setting name, in table order. Used by the settings form to build its
' list without hardcoding the keys -- add a row to tblConstants and it shows up.
' Returns a 0-length array if storage is unavailable.
Public Function ConstantKeys() As Variant
    On Error GoTo NoKeys
    ConstantKeys = TableColumnList(ConstantsTable(), CONST_KEY_COLUMN)
    Exit Function
NoKeys:
    ConstantKeys = Array()
End Function

' ---- Typed readers. Each returns the default on missing / blank / wrong type.
' The IsNumeric / IsDate guards are what stop a typo on the sheet ("fifteen")
' from raising a type-mismatch inside an unrelated macro.

Public Function ConstDouble(ByVal key As String, ByVal fallback As Double) As Double
    Dim v As Variant
    v = ConstantValue(key)
    If IsEmpty(v) Then
        ConstDouble = fallback
    ElseIf IsNumeric(v) Then
        ConstDouble = CDbl(v)
    Else
        ConstDouble = fallback
    End If
End Function

Public Function ConstDate(ByVal key As String, ByVal fallback As Date) As Date
    Dim v As Variant
    v = ConstantValue(key)
    If IsEmpty(v) Then
        ConstDate = fallback
    ElseIf IsDate(v) Then
        ConstDate = CDate(v)
    ElseIf IsNumeric(v) Then
        ' A date cell formatted as General reads back as a serial number.
        ' Guard the range so a stray 100 doesn't become 4/9/1900.
        If CDbl(v) >= 1 And CDbl(v) <= 2958465 Then     ' 2958465 = 12/31/9999
            ConstDate = CDate(CDbl(v))
        Else
            ConstDate = fallback
        End If
    Else
        ConstDate = fallback
    End If
End Function

Public Function ConstString(ByVal key As String, ByVal fallback As String) As String
    Dim v As Variant
    v = ConstantValue(key)
    If IsEmpty(v) Then
        ConstString = fallback
    Else
        ConstString = CStr(v)
    End If
End Function

' ============================================================================
'  The settings themselves.
'
'  Named exactly like the Const declarations they replaced, so no call site
'  needed editing. Read at RUN time now, not compile time -- which is the
'  entire point, but also means a value can change mid-session.
' ============================================================================

Public Property Get PERIOD_END() As Date
    PERIOD_END = ConstDate("PERIOD_END", DEF_PERIOD_END)
End Property

Public Property Get ITD_START() As Date
    ITD_START = ConstDate("ITD_START", DEF_ITD_START)
End Property

Public Property Get FORMULABAR_HEIGHT_SM() As Double
    FORMULABAR_HEIGHT_SM = ConstDouble("FORMULABAR_HEIGHT_SM", DEF_FORMULABAR_HEIGHT_SM)
End Property

Public Property Get FORMULABAR_HEIGHT_LG() As Double
    FORMULABAR_HEIGHT_LG = ConstDouble("FORMULABAR_HEIGHT_LG", DEF_FORMULABAR_HEIGHT_LG)
End Property

Public Property Get LOWER_BOUND() As Double
    LOWER_BOUND = ConstDouble("LOWER_BOUND", DEF_LOWER_BOUND)
End Property

Public Property Get UPPER_BOUND() As Double
    UPPER_BOUND = ConstDouble("UPPER_BOUND", DEF_UPPER_BOUND)
End Property

Public Property Get HEADER_ROW_HEIGHT_SINGLE() As Double
    HEADER_ROW_HEIGHT_SINGLE = ConstDouble("HEADER_ROW_HEIGHT_SINGLE", DEF_HEADER_ROW_HEIGHT_SINGLE)
End Property

Public Property Get HEADER_ROW_HEIGHT_MULTI() As Double
    HEADER_ROW_HEIGHT_MULTI = ConstDouble("HEADER_ROW_HEIGHT_MULTI", DEF_HEADER_ROW_HEIGHT_MULTI)
End Property

Public Property Get DEFAULT_ROW_HEIGHT() As Double
    DEFAULT_ROW_HEIGHT = ConstDouble("DEFAULT_ROW_HEIGHT", DEF_DEFAULT_ROW_HEIGHT)
End Property

Public Property Get DEFAULT_COLUMN_WIDTH() As Double
    DEFAULT_COLUMN_WIDTH = ConstDouble("DEFAULT_COLUMN_WIDTH", DEF_DEFAULT_COLUMN_WIDTH)
End Property

' Point size for the page footer. Separated from the footer text so the user
' can edit the wording without having to know about Excel's "&" codes.
Public Property Get FOOTER_FONT_SIZE() As Double
    FOOTER_FONT_SIZE = ConstDouble("FOOTER_FONT_SIZE", DEF_FOOTER_FONT_SIZE)
End Property

' The confidentiality footer, ASSEMBLED and ready to assign to a
' PageSetup.LeftFooter / CenterFooter / RightFooter.
'
' "&8" is not decoration -- in an Excel header/footer string, & introduces a
' formatting code and &<number> sets the point size. tblConstants stores only
' the readable prose; the size code is prepended here so a user editing the
' wording cannot delete it by accident.
Public Property Get myFooter() As String
    myFooter = "&" & CStr(CLng(FOOTER_FONT_SIZE)) & _
               ConstString("myFooter", DEF_MYFOOTER)
End Property

' The footer prose WITHOUT the size code -- what the settings form shows and
' edits. Keep these two in sync if the storage key ever changes.
Public Property Get FooterTextRaw() As String
    FooterTextRaw = ConstString("myFooter", DEF_MYFOOTER)
End Property

' ============================================================================
'  Writers -- used by the settings form. Nothing else should write here.
'
'  All of these write into ThisWorkbook (the .XLAM itself). See SaveStorage
'  for why an explicit save is not optional.
' ============================================================================

' Write one setting's Value cell, matched on the key column. Returns False if
' the key isn't in the table -- the caller decides whether that's an error.
' Does NOT save; batch your writes then call SaveStorage once.
Public Function SetConstant(ByVal key As String, ByVal newValue As Variant) As Boolean
    On Error GoTo failed
    mLastError = ""

    Dim lo As ListObject
    Set lo = ConstantsTable()
    If lo.DataBodyRange Is Nothing Then
        mLastError = "tblConstants has no rows."
        Exit Function
    End If

    Dim keyCol As Range, valCol As Range
    Set keyCol = lo.ListColumns(CONST_KEY_COLUMN).DataBodyRange
    Set valCol = lo.ListColumns(CONST_VAL_COLUMN).DataBodyRange

    Dim i As Long
    For i = 1 To keyCol.rows.Count
        If StrComp(Trim$(CStr(keyCol.Cells(i, 1).value)), key, vbTextCompare) = 0 Then
            valCol.Cells(i, 1).value = newValue
            InvalidateCache
            SetConstant = True
            Exit Function
        End If
    Next i

    mLastError = "Key '" & key & "' was not found in tblConstants."
    Exit Function

failed:
    mLastError = "Error " & Err.Number & " writing '" & key & "': " & Err.description
End Function

' Replace the whole company list with the supplied 1-D array.
'
' Sizing is done with ListObject.Resize -- NOT with ListRows.Add / .Delete.
'
' Why it matters here: Add and Delete SHIFT CELLS within the table's own
' columns. tblCompany occupies column B, and tblAlonzoChurch sits directly
' below it using column B as well. Shifting B would drag tblAlonzoChurch's
' first column out of line with its other three, so Excel refuses outright:
' "The operation is attempting to shift cells in a table on your worksheet."
'
' Resize simply redeclares where the table starts and ends. No cell ever moves,
' so neighbouring tables -- beside it or below it -- are untouched.
Public Function SetCompanyList(ByVal names As Variant) As Boolean
    On Error GoTo failed
    mLastError = ""

    Dim lo As ListObject
    Set lo = CompanyTable()

    Dim n As Long
    If IsArray(names) Then n = UBound(names) - LBound(names) + 1

    ' Blank the current contents BEFORE resizing: shrinking the table leaves
    ' any rows below the new boundary outside it, and those stale values would
    ' sit on the sheet looking like data.
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents

    ' A ListObject must always keep at least one body row, even an empty one.
    Dim rowsWanted As Long
    rowsWanted = n
    If rowsWanted < 1 Then rowsWanted = 1

    Dim newRange As Range
    Set newRange = lo.Range.Cells(1, 1).Resize(rowsWanted + 1, lo.ListColumns.Count)

    ' Refuse to grow into a neighbour rather than corrupt the sheet.
    Dim clash As String
    clash = OverlapsAnotherTable(lo, newRange)
    If Len(clash) > 0 Then
        mLastError = "The company list needs " & n & " rows, but growing the table to " & _
                     newRange.Address(False, False) & " would run into '" & clash & "'." & _
                     " Remove some names, or move that table further down the sheet."
        Exit Function
    End If

    lo.Resize newRange

    If n = 0 Then
        InvalidateCache
        SetCompanyList = True
        Exit Function
    End If

    ' Write the column in one COM call.
    Dim buf() As Variant
    ReDim buf(1 To n, 1 To 1)
    Dim i As Long, src As Long
    src = LBound(names)
    For i = 1 To n
        buf(i, 1) = names(src)
        src = src + 1
    Next i

    lo.ListColumns(COMPANY_COLUMN).DataBodyRange.value = buf

    InvalidateCache
    SetCompanyList = True
    Exit Function

failed:
    mLastError = "Error " & Err.Number & " writing the company list: " & Err.description
End Function

' Name of the first OTHER table on the same sheet that target would overlap,
' or "" if the range is clear. Used to stop a resize from eating a neighbour.
Private Function OverlapsAnotherTable(ByVal lo As ListObject, ByVal target As Range) As String
    Dim other As ListObject
    For Each other In lo.Parent.ListObjects
        If StrComp(other.name, lo.name, vbTextCompare) <> 0 Then
            If Not Application.Intersect(target, other.Range) Is Nothing Then
                OverlapsAnotherTable = other.name
                Exit Function
            End If
        End If
    Next other
End Function

' ============================================================================
'  LAMBDA LIBRARY -- every read and write of tblAlonzoChurch
'
'  The shape passed in and out of here is always the same:
'      Array(name, formula, description)
'  collected in a Collection. modLambdaLib produces that shape from a Gist, a
'  file, or a workbook's Name Manager; this module is the only thing that knows
'  it ends up in columns B/C/D of a hidden sheet.
'
'  EVERY WRITER FOLLOWS ONE PATTERN: read the whole table into memory, change
'  it there, write the whole block back. That is deliberate.
'
'  The obvious alternative -- ListRows.Add / ListRows.Delete -- SHIFTS CELLS
'  inside the table's own columns. tblAlonzoChurch shares column B with
'  tblCompany and columns D/E with tblConstants, and Excel flatly refuses a
'  shift that would disturb a neighbouring table ("The operation is attempting
'  to shift cells in a table on your worksheet"). Rewriting the block uses
'  ListObject.Resize, which redeclares the table's boundary without moving a
'  single cell, so neighbours are never at risk. It is also far fewer COM calls:
'  three array writes instead of three per row.
' ============================================================================

' Read the whole table in one COM call.
' Returns the row count; fills data (2-D, 1-based) and the three column indexes.
' Returns 0 for an empty table or a table holding only blank starter rows.
'
' .Value, not .Value2, for the same reason as LoadCache: .Value2 hands back a
' date-looking description as a serial number, and "43466" is not what the user
' typed. The table always has three columns, so .Value is always a 2-D array --
' the single-cell-returns-a-scalar trap cannot bite here.
Private Function ReadLambdaBlock(ByRef data As Variant, ByRef ixName As Long, _
                                 ByRef ixCode As Long, ByRef ixDesc As Long) As Long
    Dim lo As ListObject
    Set lo = LambdaTable()

    ixName = lo.ListColumns(LAMBDA_NAME_COLUMN).index
    ixCode = lo.ListColumns(LAMBDA_CODE_COLUMN).index
    ixDesc = lo.ListColumns(LAMBDA_DESC_COLUMN).index

    If lo.DataBodyRange Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Function

    data = lo.DataBodyRange.value
    ReadLambdaBlock = UBound(data, 1)
End Function

' Every stored LAMBDA as Array(name, formula, description), in table order.
' Rows with a blank Function name are skipped. Never returns Nothing.
Public Function LambdaAll() As Collection
    Dim out As New Collection
    Set LambdaAll = out

    On Error GoTo done
    Dim data As Variant, ixName As Long, ixCode As Long, ixDesc As Long
    Dim n As Long, i As Long, nm As String
    n = ReadLambdaBlock(data, ixName, ixCode, ixDesc)

    For i = 1 To n
        nm = Trim$(CStr(data(i, ixName) & ""))
        If Len(nm) > 0 Then
            out.Add Array(nm, CStr(data(i, ixCode) & ""), CStr(data(i, ixDesc) & ""))
        End If
    Next i

done:
End Function

' Just the function names, 1-based. Returns a 0-length array if empty --
' guard with HasItems(), exactly like CompanyList().
Public Function LambdaNames() As Variant
    On Error GoTo NoNames
    LambdaNames = TableColumnList(LambdaTable(), LAMBDA_NAME_COLUMN)
    Exit Function
NoNames:
    LambdaNames = Array()
End Function

' How many LAMBDAs are stored.
Public Function LambdaCount() As Long
    LambdaCount = LambdaAll().Count
End Function

' 1-based position of a function in LambdaAll(), or 0 if it isn't there.
' Case-insensitive: Excel defined names are, so the library must be too.
Public Function LambdaRowIndex(ByVal name As String) As Long
    Dim all As Collection, i As Long
    Set all = LambdaAll()
    For i = 1 To all.Count
        If StrComp(CStr(all(i)(0)), name, vbTextCompare) = 0 Then
            LambdaRowIndex = i
            Exit Function
        End If
    Next i
End Function

' One entry as Array(name, formula, description), or Empty if not found.
Public Function LambdaEntry(ByVal name As String) As Variant
    Dim all As Collection, i As Long
    Set all = LambdaAll()
    For i = 1 To all.Count
        If StrComp(CStr(all(i)(0)), name, vbTextCompare) = 0 Then
            LambdaEntry = all(i)
            Exit Function
        End If
    Next i
End Function

' True if the library already holds a function by this name.
Public Function LambdaExists(ByVal name As String) As Boolean
    LambdaExists = Not IsEmpty(LambdaEntry(name))
End Function

Public Function LambdaFormula(ByVal name As String) As String
    Dim e As Variant
    e = LambdaEntry(name)
    If Not IsEmpty(e) Then LambdaFormula = CStr(e(1))
End Function

Public Function LambdaDescription(ByVal name As String) As String
    Dim e As Variant
    e = LambdaEntry(name)
    If Not IsEmpty(e) Then LambdaDescription = CStr(e(2))
End Function

' ----------------------------------------------------------------------------
'  Writers. All of them funnel into WriteLambdaBlock.
'  None of them save -- batch your changes, then call SaveStorage once.
' ----------------------------------------------------------------------------

' Replace the entire table with the supplied entries. The single choke point
' every other writer goes through.
Public Function WriteLambdaBlock(ByVal entries As Collection) As Boolean
    On Error GoTo failed
    mLastError = ""

    Dim lo As ListObject
    Set lo = LambdaTable()

    Dim n As Long
    n = entries.Count

    ' Blank the current contents BEFORE resizing. Shrinking leaves any rows
    ' below the new boundary outside the table, and those stale values would sit
    ' on the sheet looking like live data.
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents

    ' A ListObject must always keep at least one body row, even an empty one.
    Dim rowsWanted As Long
    rowsWanted = n
    If rowsWanted < 1 Then rowsWanted = 1

    ' A totals row would be counted into the resize and end up as data.
    Dim hadTotals As Boolean
    hadTotals = lo.ShowTotals
    If hadTotals Then lo.ShowTotals = False

    Dim newRange As Range
    Set newRange = lo.Range.Cells(1, 1).Resize(rowsWanted + 1, lo.ListColumns.Count)

    ' Refuse to grow into a neighbour rather than corrupt the sheet.
    Dim clash As String
    clash = OverlapsAnotherTable(lo, newRange)
    If Len(clash) > 0 Then
        If hadTotals Then lo.ShowTotals = True
        mLastError = "The library needs " & n & " rows, but growing the table to " & _
                     newRange.Address(False, False) & " would run into '" & clash & "'." & _
                     " Move that table further down the sheet, or remove some functions."
        Exit Function
    End If

    lo.Resize newRange

    If n = 0 Then
        If hadTotals Then lo.ShowTotals = True
        WriteLambdaBlock = True
        Exit Function
    End If

    ' Build the three columns, then write each in ONE COM call.
    Dim nameArr() As Variant, codeArr() As Variant, descArr() As Variant
    ReDim nameArr(1 To n, 1 To 1)
    ReDim codeArr(1 To n, 1 To 1)
    ReDim descArr(1 To n, 1 To 1)

    Dim i As Long, e As Variant
    For i = 1 To n
        e = entries(i)
        nameArr(i, 1) = e(0)
        codeArr(i, 1) = e(1)
        descArr(i, 1) = e(2)
    Next i

    ' Text format FIRST, then the values. The other order lets Excel try to
    ' evaluate "=LAMBDA(...)" as a formula on the way in, which fails.
    With lo.ListColumns(LAMBDA_CODE_COLUMN).DataBodyRange
        .NumberFormat = "@"
        .value = codeArr
        .WrapText = True
    End With
    lo.ListColumns(LAMBDA_NAME_COLUMN).DataBodyRange.value = nameArr
    With lo.ListColumns(LAMBDA_DESC_COLUMN).DataBodyRange
        .value = descArr
        .WrapText = True
    End With

    If hadTotals Then lo.ShowTotals = True

    WriteLambdaBlock = True
    Exit Function

failed:
    mLastError = "Error " & Err.Number & " writing the LAMBDA library: " & Err.description
End Function

' Add a function, or overwrite it if the name is already taken.
Public Function UpsertLambda(ByVal name As String, ByVal formula As String, _
                             ByVal description As String) As Boolean
    Dim one As New Collection
    one.Add Array(name, formula, description)

    Dim added As Long, updated As Long
    UpsertLambda = UpsertLambdas(one, added, updated)
End Function

' Bulk add-or-overwrite. Reports how many were new vs. replaced so the caller
' can tell the user something more useful than "done".
'
' Bulk matters: the import paths add dozens of functions at once, and doing that
' one UpsertLambda at a time would resize and rewrite the whole table per entry.
Public Function UpsertLambdas(ByVal entries As Collection, _
                              ByRef addedCount As Long, _
                              ByRef updatedCount As Long) As Boolean
    On Error GoTo failed
    mLastError = ""
    addedCount = 0
    updatedCount = 0

    Dim current As Collection
    Set current = LambdaAll()

    ' Nothing on either side means nothing to do. This guard is not cosmetic:
    ' ReDim buf(1 To 0) further down is a runtime error, not an empty array.
    If current.Count = 0 And entries.Count = 0 Then
        UpsertLambdas = True
        Exit Function
    End If

    ' name -> 1-based position in current, for an O(1) hit instead of a rescan
    ' per incoming entry.
    Dim pos As Object
    Set pos = CreateObject("Scripting.Dictionary")
    pos.CompareMode = 1                              ' 1 = vbTextCompare
    Dim i As Long
    For i = 1 To current.Count
        pos(CStr(current(i)(0))) = i
    Next i

    ' A Collection can't be updated in place, so build the result as an array.
    Dim buf() As Variant
    ReDim buf(1 To current.Count + entries.Count)
    Dim used As Long
    used = current.Count
    For i = 1 To current.Count
        buf(i) = current(i)
    Next i

    Dim e As Variant, nm As String
    For Each e In entries
        nm = Trim$(CStr(e(0)))
        If Len(nm) > 0 Then
            If pos.Exists(nm) Then
                ' Keep the EXISTING name's capitalisation? No -- take the
                ' incoming one, so a deliberate re-import can fix a typo.
                buf(pos(nm)) = Array(nm, CStr(e(1)), CStr(e(2)))
                updatedCount = updatedCount + 1
            Else
                used = used + 1
                buf(used) = Array(nm, CStr(e(1)), CStr(e(2)))
                pos(nm) = used
                addedCount = addedCount + 1
            End If
        End If
    Next e

    Dim out As New Collection
    For i = 1 To used
        out.Add buf(i)
    Next i

    UpsertLambdas = WriteLambdaBlock(out)
    Exit Function

failed:
    mLastError = "Error " & Err.Number & " updating the LAMBDA library: " & Err.description
End Function

' Remove one function. Returns False (with LastError set) if it isn't there.
Public Function DeleteLambda(ByVal name As String) As Boolean
    On Error GoTo failed
    mLastError = ""

    Dim current As Collection, out As New Collection
    Set current = LambdaAll()

    Dim i As Long, found As Boolean
    For i = 1 To current.Count
        If StrComp(CStr(current(i)(0)), name, vbTextCompare) = 0 Then
            found = True
        Else
            out.Add current(i)
        End If
    Next i

    If Not found Then
        mLastError = "'" & name & "' is not in the LAMBDA library."
        Exit Function
    End If

    DeleteLambda = WriteLambdaBlock(out)
    Exit Function

failed:
    mLastError = "Error " & Err.Number & " deleting '" & name & "': " & Err.description
End Function

' Change a function's name, keeping its formula and description.
' Refuses if the new name is already taken by a DIFFERENT function.
Public Function RenameLambda(ByVal oldName As String, ByVal newName As String) As Boolean
    On Error GoTo failed
    mLastError = ""

    Dim current As Collection, out As New Collection
    Set current = LambdaAll()

    Dim i As Long, found As Boolean
    For i = 1 To current.Count
        If StrComp(CStr(current(i)(0)), newName, vbTextCompare) = 0 And _
           StrComp(CStr(current(i)(0)), oldName, vbTextCompare) <> 0 Then
            mLastError = "A function named '" & newName & "' is already in the library."
            Exit Function
        End If
    Next i

    For i = 1 To current.Count
        If StrComp(CStr(current(i)(0)), oldName, vbTextCompare) = 0 Then
            found = True
            out.Add Array(newName, current(i)(1), current(i)(2))
        Else
            out.Add current(i)
        End If
    Next i

    If Not found Then
        mLastError = "'" & oldName & "' is not in the LAMBDA library."
        Exit Function
    End If

    RenameLambda = WriteLambdaBlock(out)
    Exit Function

failed:
    mLastError = "Error " & Err.Number & " renaming '" & oldName & "': " & Err.description
End Function

' Change a function's description, leaving name and formula alone.
Public Function SetLambdaDescription(ByVal name As String, ByVal description As String) As Boolean
    On Error GoTo failed
    mLastError = ""

    Dim current As Collection, out As New Collection
    Set current = LambdaAll()

    Dim i As Long, found As Boolean
    For i = 1 To current.Count
        If StrComp(CStr(current(i)(0)), name, vbTextCompare) = 0 Then
            found = True
            out.Add Array(current(i)(0), current(i)(1), description)
        Else
            out.Add current(i)
        End If
    Next i

    If Not found Then
        mLastError = "'" & name & "' is not in the LAMBDA library."
        Exit Function
    End If

    SetLambdaDescription = WriteLambdaBlock(out)
    Exit Function

failed:
    mLastError = "Error " & Err.Number & " describing '" & name & "': " & Err.description
End Function

' Persist the add-in to disk.
'
' NOT optional. Writing to a loaded .XLAM marks it dirty, but Excel closes
' add-ins at shutdown WITHOUT prompting to save. Skip this and the user's
' settings vanish at the end of the session with no warning at all.
Public Sub SaveStorage()
    On Error Resume Next
    ThisWorkbook.Save
End Sub
