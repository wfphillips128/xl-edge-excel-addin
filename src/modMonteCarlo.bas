Attribute VB_Name = "modMonteCarlo"
' =============================================================================
' modMonteCarlo - Monte Carlo LAMBDA library. All logic, no user interface.
'
' Layering, the same three tiers the LAMBDA Studio feature uses:
'
'   modMonteCarloRibbon   every MsgBox and InputBox lives there
'   modMonteCarlo         this module - no UI, so it is drivable from the
'                         Immediate window:  ? McFindDistribution("pert")
'   modLambdaLib          reused as-is for fetching and installing definitions
'
' WHERE THE FUNCTIONS COME FROM. A copy of the library is BUNDLED in the add-in,
' on the very-hidden sheet shtMonteCarloLib, pinned to MC_LIBRARY_VERSION. Each
' ribbon action installs into the user's workbook only the functions it needs,
' plus everything those call - three names for a PERT, not all seventy-odd. No
' network: the bundle is the source, so installing works offline and the VBA
' and the LAMBDAs can never be out of step. The gist remains the public copy for
' people without XL Edge.
'
' WHAT THIS FEATURE ACTUALLY DOES. It writes a native formula into a cell:
'
'   =fx.RiskPert<lambda>($C$4, $C$5, $C$6, MC_Trials, 7)
'
' and nothing more. The formula calls LAMBDA functions stored as defined names
' in the user's own workbook, so the finished file keeps calculating for people
' who have never heard of XL Edge. The add-in is an AUTHORING tool. If it were
' ever uninstalled, every model built with it would carry on working.
'
' That is the whole point of the exercise - the add-ins this replaces all fail
' with #NAME? the moment the recipient does not have them installed.
'
' THE LAMBDA CHARACTER. Every function name ends in U+03BB (Greek lambda). It is NOT
' written literally anywhere in this module. A .bas file exports as Windows-1252
' and it has no representation in that code page, so a literal would be silently
' mangled on the round trip through export and import. Lam() builds it with
' ChrW instead, which survives anything.
' =============================================================================
Option Explicit

' The published library. A single-file gist, which is a hard requirement:
' NormalizeGistUrl rewrites to .../raw, and that returns only the FIRST file.
Public Const MC_GIST_URL As String = "https://gist.github.com/wfphillips128/f91bff77212ab2c3d8f55a4f0a51b8b6"

Public Const MC_TRIALS_NAME As String = "MC_Trials"
Public Const MC_VARID_NAME As String = "MC_NextVarID"
Public Const MC_DEF_TRIALS As Long = 10000

' Rows in each result block, as the LAMBDAs build them with their defaults.
' The ribbon states these up front so the user can leave room and avoid
' #SPILL!. If a LAMBDA's default layout changes, change the number here too.
Public Const MC_STATS_ROWS As Long = 12          ' fx.RiskStats: name + 8 statistics + P10/P50/P90
Public Const MC_STATS_DETAIL_ROWS As Long = 33   ' fx.RiskStatsDetail: name + 13 statistics + P5..P95
Public Const MC_HIST_ROWS As Long = 21           ' fx.RiskHist: header + default 20 bins
Public Const MC_RISK_ROWS As Long = 5            ' fx.RiskMeasures: name, Confidence, VaR, CVaR, ES

' The chart-data blocks. Charts point at plain ranges inside these spills, so
' the sizes are FIXED - the LAMBDAs' defaults - and must match them.
' fx.RiskChartHist is ONE spill: title, P-line header, P10/P50/P90, bin header,
' then the bins - MC_CHART_HEAD_ROWS + bins rows, MC_CHART_HIST_COLS wide.
' The rows above the bins never move, whatever the bin count.
Public Const MC_CHART_BINS As Long = 30          ' fx.RiskChartHist default bins
Public Const MC_CHART_MARKS As Long = 3          ' P10, P50, P90
Public Const MC_CHART_HEAD_ROWS As Long = 6      ' 1 title + 1 header + 3 P-lines + 1 header
Public Const MC_CHART_HIST_COLS As Long = 8
Public Const MC_CHART_TITLE_ROWS As Long = 1     ' title row atop fx.RiskTornado

' The library version this VBA is written against. The chart code's fixed
' block sizes, the catalog's argument lists and the function names all depend
' on it, so the bundle must match EXACTLY: McLoadBundleFromFile refuses any
' other MonteCarlo.txt, and the install layer refuses a stale bundle.
Public Const MC_LIBRARY_VERSION As String = "0.6.0"

' Hidden name stamped into each workbook: the library version it holds.
Public Const MC_VERSION_NAME As String = "MC_LibraryVersion"

' Where the developer build step reads the library from by default.
Public Const MC_SOURCE_PATH As String = "Z:\ACCELERATOR\Monte Carlo software\lambda\MonteCarlo.txt"

' The bundled copy: Function | Formula | Description from row 2, the version in
' F1. Owned by this module alone - AddInStorage keeps XL Edge's own tables, and
' this is a separate sheet so a resize there can never touch it.
Private Const BUNDLE_SHEET As String = "shtMonteCarloLib"

' Rotation is the LAST argument of the Archimedean copulas, after Trials, VarID
' and Seed, so formulas written before it existed keep working. "?^" marks a
' parameter as optional and trailing; McBuildCall puts it after the tail.
Private Const ROTATION_PARAM As String = _
    "?^Rotation: 0, 90, 180 or 270  (default 0. 180 moves the tail to the other end; 90 and 270 make the dependence negative, 2 variables only)"
Private Const BUNDLE_VERSION_CELL As String = "F1"

Private mCatalog As Collection
Private mCopulas As Collection
Private mBundle As Object               ' name -> Array(name, formula, description)
Private mBundleOrder As Collection      ' names, in library order


' --- the lambda suffix -------------------------------------------------------

' U+03BB GREEK SMALL LETTER LAMBDA. See the note in the module header for why
' this is ChrW and not a literal.
Public Function Lam() As String
    Lam = ChrW$(&H3BB)
End Function

Public Function McFunctionName(ByVal baseName As String) As String
    McFunctionName = baseName & Lam()
End Function


' --- the distribution catalog ---------------------------------------------

' Each entry: Array(display name, function base name, pipe-delimited parameters,
'                   ribbon item id)
'
' Parameters are listed in the order the LAMBDA takes them, and every optional
' parameter that comes BEFORE Trials is listed too - Beta's bounds, for example.
' Skipping one would shift Trials into the wrong position.
'
' A leading "?" marks a parameter as optional: the prompt says so, and typing
' "none" leaves it out of the formula (McIsNone), so the LAMBDA's own default
' applies. The "?" is never shown to the user.
'
' The ribbon item id must match an <item id="..."> in the Continuous or
' Discrete gallery of ribbon - Monte Carlo menu items.xml. That is how a click
' on an icon finds its way back to this list.
Public Function McCatalog() As Collection
    If Not mCatalog Is Nothing Then
        Set McCatalog = mCatalog
        Exit Function
    End If

    Dim c As New Collection
    c.Add Array("Uniform", "fx.RiskUniform", "Min|Max", "McD_Uniform")
    c.Add Array("Normal", "fx.RiskNormal", "Mean|Standard deviation", "McD_Normal")
    c.Add Array("Lognormal", "fx.RiskLogNorm", "Mean of ln(X)|Standard deviation of ln(X)", "McD_Lognormal")
    c.Add Array("Triangular", "fx.RiskTriang", "Min|Most likely|Max", "McD_Triangular")
    c.Add Array("PERT", "fx.RiskPert", "Min|Most likely|Max", "McD_Pert")
    c.Add Array("Beta", "fx.RiskBeta", "Shape1 (alpha)|Shape2 (beta)|?Lower bound (default 0)|?Upper bound (default 1)", "McD_Beta")
    c.Add Array("Gamma", "fx.RiskGamma", "Alpha (shape)|Beta (scale)", "McD_Gamma")
    c.Add Array("Erlang", "fx.RiskErlang", "K (whole number)|Beta (scale)", "McD_Erlang")
    c.Add Array("Exponential", "fx.RiskExpon", "Mean (not the rate)", "McD_Exponential")
    c.Add Array("Weibull", "fx.RiskWeibull", "Alpha (shape)|Beta (scale)", "McD_Weibull")
    c.Add Array("Bernoulli", "fx.RiskBernoulli", "P", "McD_Bernoulli")
    c.Add Array("Binomial", "fx.RiskBinomial", "N (trials)|P", "McD_Binomial")
    c.Add Array("Discrete", "fx.RiskDiscrete", "Values|Probabilities (must sum to 1)", "McD_Discrete")
    c.Add Array("Discrete uniform", "fx.RiskDUniform", "Values", "McD_DUniform")
    c.Add Array("Cumulative", "fx.RiskCumul", "Lowest value|Highest value|X values|Y probabilities", "McD_Cumulative")

    ' v0.3 continuous
    c.Add Array("Cauchy", "fx.RiskCauchy", "Location (median)|Scale", "McD_Cauchy")
    c.Add Array("Chi-squared", "fx.RiskChiSq", "Degrees of freedom", "McD_ChiSq")
    c.Add Array("F", "fx.RiskF", "Numerator degrees of freedom|Denominator degrees of freedom", "McD_F")
    c.Add Array("Gumbel", "fx.RiskExtValue", "Location (mode)|Scale", "McD_Gumbel")
    c.Add Array("Half-Cauchy", "fx.RiskHalfCauchy", "Scale|?Location (lower bound, default 0)", "McD_HalfCauchy")
    c.Add Array("Half-normal", "fx.RiskHalfNormal", "Scale|?Location (lower bound, default 0)", "McD_HalfNormal")
    c.Add Array("Half-Student t", "fx.RiskHalfStudent", "Degrees of freedom|Scale|?Location (lower bound, default 0)", "McD_HalfStudent")
    c.Add Array("Inverse chi-squared", "fx.RiskInvChiSq", "Degrees of freedom", "McD_InvChiSq")
    c.Add Array("Inverse gamma", "fx.RiskInvGamma", "Alpha (shape)|Beta (scale)", "McD_InvGamma")
    c.Add Array("Inverse Gaussian", "fx.RiskInvGauss", "Mean|Shape (lambda)", "McD_InvGauss")
    c.Add Array("Laplace", "fx.RiskLaplace", "Location|Scale (not the standard deviation)", "McD_Laplace")
    c.Add Array("Logistic", "fx.RiskLogistic", "Location|Scale (not the standard deviation)", "McD_Logistic")
    c.Add Array("Noncentral beta", "fx.RiskNCBeta", "Shape1|Shape2|Noncentrality", "McD_NCBeta")
    c.Add Array("Noncentral F", "fx.RiskNCF", "Numerator degrees of freedom|Denominator degrees of freedom|Noncentrality", "McD_NCF")
    c.Add Array("Noncentral t", "fx.RiskNCStudent", "Degrees of freedom|Noncentrality", "McD_NCStudent")
    c.Add Array("Pareto", "fx.RiskPareto", "Theta (shape)|A (scale - the minimum)", "McD_Pareto")
    c.Add Array("Skew normal", "fx.RiskSkewNormal", "Location|Scale|Shape (skewness; 0 = normal)", "McD_SkewNormal")
    c.Add Array("Student t", "fx.RiskStudent", "Degrees of freedom", "McD_Student")
    c.Add Array("Truncated normal", "fx.RiskTruncNormal", "Mean|Standard deviation|?Lower bound|?Upper bound", "McD_TruncNormal")

    ' v0.3 discrete
    c.Add Array("Benford", "fx.RiskBenford", "?Digits (1 = first digit, 2 = first two)", "McD_Benford")
    c.Add Array("Geometric", "fx.RiskGeomet", "P (counts failures before the first success)", "McD_Geometric")
    c.Add Array("Hypergeometric", "fx.RiskHypergeo", "Draws (n)|Successes in the population (D)|Population size (M)", "McD_Hypergeo")
    c.Add Array("Negative binomial", "fx.RiskNegbin", "Successes (s)|P (counts failures before the s-th success)", "McD_Negbin")
    c.Add Array("Poisson", "fx.RiskPoisson", "Mean", "McD_Poisson")

    Set mCatalog = c
    Set McCatalog = c
End Function

Public Function McCatalogCount() As Long
    McCatalogCount = McCatalog().count
End Function

Public Function McCatalogItem(ByVal idx As Long) As Variant
    If idx < 1 Or idx > McCatalogCount() Then Exit Function
    McCatalogItem = McCatalog().item(idx)
End Function

' --- the copula catalog ----------------------------------------------------

' Same shape as McCatalog. The ribbon item ids are the <button id="McC_...">
' entries in the Insert Monte Carlo Copula menu. A copula spills Trials rows by
' one column per variable, so it takes the same Trials and VarID tail as a
' distribution and McBuildCall writes it unchanged.
Public Function McCopulaCatalog() As Collection
    If Not mCopulas Is Nothing Then
        Set McCopulaCatalog = mCopulas
        Exit Function
    End If

    Dim c As New Collection
    c.Add Array("Gaussian copula", "fx.RiskCopulaGauss", _
                "Correlation matrix (select the square range, 1s on the diagonal)", "McC_Gauss")
    c.Add Array("Student t copula", "fx.RiskCopulaT", _
                "Correlation matrix (select the square range, 1s on the diagonal)|Degrees of freedom (smaller = stronger tail dependence)", "McC_T")
    c.Add Array("Clayton copula", "fx.RiskCopulaClayton", _
                "Theta, > 0  (Kendall's tau = Theta / (Theta + 2))|?Number of variables (default 2)|" & ROTATION_PARAM, "McC_Clayton")
    c.Add Array("Gumbel copula", "fx.RiskCopulaGumbel", _
                "Theta, >= 1  (Kendall's tau = 1 - 1/Theta)|?Number of variables (default 2)|" & ROTATION_PARAM, "McC_Gumbel")
    c.Add Array("Frank copula", "fx.RiskCopulaFrank", _
                "Theta, > 0  (larger = stronger)|?Number of variables (default 2)|" & ROTATION_PARAM, "McC_Frank")
    ' The negative-dependence buttons: same functions, two variables (the
    ' default), Theta below 0.
    c.Add Array("Clayton copula, negative dependence", "fx.RiskCopulaClayton", _
                "Theta, between -1 and 0  (Kendall's tau = Theta / (Theta + 2); -1 = perfectly opposite)", "McC_ClaytonNeg")
    c.Add Array("Frank copula, negative dependence", "fx.RiskCopulaFrank", _
                "Theta, below 0  (more negative = stronger)", "McC_FrankNeg")

    Set mCopulas = c
    Set McCopulaCatalog = c
End Function

' Returns the catalog entry for a ribbon button id, or Empty when unknown.
Public Function McCopulaByItemId(ByVal itemId As String) As Variant
    Dim entry As Variant
    For Each entry In McCopulaCatalog()
        If StrComp(CStr(entry(3)), itemId, vbTextCompare) = 0 Then
            McCopulaByItemId = entry
            Exit Function
        End If
    Next entry
End Function

' How many columns a copula call will spill, from the arguments the user gave:
' the width of the correlation matrix, or the Dimensions argument (default 2).
' Returns 0 when it cannot tell - a typed-in array constant, say - and the
' caller then describes the size in words instead.
Public Function McCopulaColumns(ByVal baseName As String, ByVal args As Variant) As Long
    Dim v As Variant
    On Error GoTo Unknown

    If baseName = "fx.RiskCopulaGauss" Or baseName = "fx.RiskCopulaT" Then
        v = Application.Evaluate("COLUMNS(" & CStr(args(0)) & ")")
    ElseIf UBound(args) < 1 Then
        v = 2
    ElseIf Len(CStr(args(1))) = 0 Then
        v = 2
    Else
        v = Application.Evaluate(CStr(args(1)))
    End If
    If IsNumeric(v) Then McCopulaColumns = CLng(v)
    Exit Function

Unknown:
    McCopulaColumns = 0
End Function


' Maps a ribbon gallery item id ("McD_Pert") to its catalog index.
' Returns 0 when the id is not in the catalog.
Public Function McFindByItemId(ByVal itemId As String) As Long
    Dim i As Long, entry As Variant
    For i = 1 To McCatalogCount()
        entry = McCatalogItem(i)
        If StrComp(CStr(entry(3)), itemId, vbTextCompare) = 0 Then
            McFindByItemId = i
            Exit Function
        End If
    Next i
End Function

' Accepts "4", "pert", "PERT" - a number first, then a case-insensitive
' partial name match. Returns 0 when nothing matches.
Public Function McFindDistribution(ByVal key As String) As Long
    Dim k As String, i As Long, entry As Variant
    k = Trim$(key)
    If Len(k) = 0 Then Exit Function

    If IsNumeric(k) Then
        If CLng(val(k)) >= 1 And CLng(val(k)) <= McCatalogCount() Then
            McFindDistribution = CLng(val(k))
        End If
        Exit Function
    End If

    For i = 1 To McCatalogCount()
        entry = McCatalogItem(i)
        If StrComp(CStr(entry(0)), k, vbTextCompare) = 0 Then
            McFindDistribution = i
            Exit Function
        End If
    Next i
    For i = 1 To McCatalogCount()
        entry = McCatalogItem(i)
        If InStr(1, CStr(entry(0)), k, vbTextCompare) > 0 Then
            McFindDistribution = i
            Exit Function
        End If
    Next i
End Function


' --- the library: a bundled copy, installed as needed ---------------------

' True when the bundle is present and matches MC_LIBRARY_VERSION; otherwise
' why says what is wrong and how to fix it.
Public Function McBundleReady(Optional ByRef why As String) As Boolean
    LoadBundle why
    McBundleReady = (Len(why) = 0)
End Function

Public Function McBundleCount() As Long
    Dim why As String
    LoadBundle why
    If Len(why) = 0 Then McBundleCount = mBundleOrder.count
End Function

' Read the bundle sheet once per session into a dictionary.
Private Sub LoadBundle(ByRef why As String)
    Dim ws As Worksheet, data As Variant, n As Long, i As Long, nm As String, ver As String
    why = ""
    If Not mBundle Is Nothing Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(BUNDLE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        why = "The Monte Carlo library is not bundled in this copy of XL Edge " & _
              "(the sheet " & BUNDLE_SHEET & " is missing). The add-in was built without " & _
              "running McDevRefreshBundle."
        Exit Sub
    End If
    ver = CStr(ws.Range(BUNDLE_VERSION_CELL).Value2)
    If ver <> MC_LIBRARY_VERSION Then
        why = "The Monte Carlo library bundled in XL Edge is version " & ver & _
              ", but this version of XL Edge needs " & MC_LIBRARY_VERSION & _
              ". Run McDevRefreshBundle and rebuild the add-in."
        Exit Sub
    End If
    n = ws.Cells(ws.rows.count, 1).End(xlUp).Row - 1
    If n < 1 Then
        why = "The Monte Carlo library bundled in XL Edge is empty. Run McDevRefreshBundle."
        Exit Sub
    End If

    data = ws.Range("A2").Resize(n, 3).Value2
    Set mBundle = CreateObject("Scripting.Dictionary")
    mBundle.CompareMode = 1                          ' vbTextCompare, as Excel names are
    Set mBundleOrder = New Collection
    For i = 1 To n
        nm = Trim$(CStr(data(i, 1)))
        If Len(nm) > 0 Then
            mBundle(nm) = Array(nm, CStr(data(i, 2)), CStr(data(i, 3)))
            mBundleOrder.Add nm
        End If
    Next i
End Sub

' Every fx.Risk...<lambda> name a formula mentions. The lambda at the end of
' every library name is what makes a plain text scan reliable: nothing else in
' a formula looks like that. Comments are already gone by the time a formula
' is stored, and a spurious match would only install one function too many.
Private Function ReferencedNames(ByVal formula As String) As Collection
    Dim out As New Collection, seen As Object
    Dim p As Long, q As Long, tok As String
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    p = InStr(1, formula, "fx.Risk", vbTextCompare)
    Do While p > 0
        q = InStr(p, formula, Lam())
        If q = 0 Then Exit Do
        tok = Mid$(formula, p, q - p + 1)
        If IsPlainWord(Mid$(tok, 4, Len(tok) - 4)) Then
            If Not seen.Exists(tok) Then
                seen(tok) = True
                out.Add tok
            End If
        End If
        p = InStr(p + 1, formula, "fx.Risk", vbTextCompare)
    Loop
    Set ReferencedNames = out
End Function

Private Function IsPlainWord(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not (ch Like "[A-Za-z0-9]") Then Exit Function
    Next i
    IsPlainWord = True
End Function

' The given functions plus everything they call, directly or not, as bundle
' entries in library order. names: an Array or Collection of base names
' ("fx.RiskPert") or full names; unknown names are ignored.
Public Function McDependencyClosure(ByVal names As Variant) As Collection
    Dim out As New Collection, want As Object, stack As New Collection
    Dim nm As Variant, dep As Variant, why As String, key As String

    Set McDependencyClosure = out
    LoadBundle why
    If Len(why) > 0 Then Exit Function

    Set want = CreateObject("Scripting.Dictionary")
    want.CompareMode = 1
    For Each nm In names
        stack.Add FullName(CStr(nm))
    Next nm
    Do While stack.count > 0
        key = stack(stack.count)
        stack.Remove stack.count
        If Not want.Exists(key) And mBundle.Exists(key) Then
            want(key) = True
            For Each dep In ReferencedNames(CStr(mBundle(key)(1)))
                stack.Add CStr(dep)
            Next dep
        End If
    Loop
    For Each nm In mBundleOrder
        If want.Exists(nm) Then out.Add mBundle(nm)
    Next nm
End Function

Private Function FullName(ByVal nm As String) As String
    If Right$(nm, 1) = Lam() Then FullName = nm Else FullName = nm & Lam()
End Function

' The library's functions currently defined in a workbook (workbook scope).
Private Function InstalledNames(ByVal wb As Workbook) As Collection
    Dim out As New Collection, nm As name, s As String
    For Each nm In wb.names
        s = nm.name
        If StrComp(Left$(s, 7), "fx.Risk", vbTextCompare) = 0 And Right$(s, 1) = Lam() Then out.Add s
    Next nm
    Set InstalledNames = out
End Function

Public Function McLibraryInstalled(ByVal wb As Workbook) As Boolean
    If wb Is Nothing Then Exit Function
    McLibraryInstalled = (InstalledNames(wb).count > 0)
End Function

' Installed, and at the version this add-in ships.
Public Function McLibraryCurrent(ByVal wb As Workbook) As Boolean
    McLibraryCurrent = McLibraryInstalled(wb) And (McWorkbookVersion(wb) = MC_LIBRARY_VERSION)
End Function

' The library version stamped in a workbook, or "" if none.
Public Function McWorkbookVersion(ByVal wb As Workbook) As String
    Dim s As String
    On Error Resume Next
    s = wb.names(MC_VERSION_NAME).RefersTo
    On Error GoTo 0
    s = Replace(Replace(s, "=", ""), """", "")
    McWorkbookVersion = Trim$(s)
End Function

Private Sub StampVersion(ByVal wb As Workbook)
    On Error Resume Next
    wb.names(MC_VERSION_NAME).Delete
    On Error GoTo 0
    wb.names.Add name:=MC_VERSION_NAME, RefersTo:="=""" & MC_LIBRARY_VERSION & """", Visible:=False
End Sub

Private Function HasName(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim s As String
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    s = wb.names(nm).RefersTo
    On Error GoTo 0
    HasName = (Len(s) > 0)
End Function

' THE ONE ENTRY POINT every ribbon action uses before writing a formula.
' Makes sure wb can evaluate the named functions, and returns True if it can.
'
'   Same version as the bundle (or a new workbook): install only what is
'   missing from the dependency closure.
'   Older, or unstamped with library names present: REFRESH every library
'   name already in the workbook as well, so the workbook never mixes versions
'   - a v0.3 fx.RiskU under a v0.5 function is exactly the bug this prevents.
'   fromVersion then says what it was, for the message.
'
' added counts the names written. why explains a False.
Public Function McEnsureFunctions(ByVal wb As Workbook, ByVal names As Variant, _
                                  ByRef added As Long, ByRef fromVersion As String, _
                                  Optional ByRef why As String) As Boolean
    Dim installed As Collection, todo As New Collection, Failed As New Collection
    Dim both As New Collection, e As Variant, nm As Variant, cur As String

    added = 0
    fromVersion = ""
    why = ""
    If wb Is Nothing Then Exit Function
    If Not McBundleReady(why) Then Exit Function

    Set installed = InstalledNames(wb)
    cur = McWorkbookVersion(wb)

    If installed.count > 0 And cur <> MC_LIBRARY_VERSION Then
        fromVersion = IIf(Len(cur) > 0, cur, "an earlier version")
        For Each nm In names
            both.Add nm
        Next nm
        For Each nm In installed
            both.Add nm
        Next nm
        Set todo = McDependencyClosure(both)
    Else
        For Each e In McDependencyClosure(names)
            If Not HasName(wb, CStr(e(0))) Then todo.Add e
        Next e
    End If

    If todo.count > 0 Then added = modLambdaLib.AddEntriesToWorkbook(todo, wb, Failed)

    For Each e In McDependencyClosure(names)
        If Not HasName(wb, CStr(e(0))) Then
            why = CStr(e(0)) & " could not be added to the workbook."
            If Failed.count > 0 Then why = why & vbCrLf & vbCrLf & "Excel said: " & CStr(Failed(1))
            Exit Function
        End If
    Next e
    If McDependencyClosure(names).count = 0 Then
        why = "None of the requested functions are in the bundled library."
        Exit Function
    End If

    If cur <> MC_LIBRARY_VERSION Then StampVersion wb
    McEnsureFunctions = True
End Function

' Every library function, into wb. Returns how many were written.
Public Function McInstallAll(ByVal wb As Workbook, _
                             Optional ByRef Failed As Collection, _
                             Optional ByRef wrapped As Collection) As Long
    Dim all As New Collection, nm As Variant, why As String
    If wb Is Nothing Then Exit Function
    LoadBundle why
    If Len(why) > 0 Then Err.Raise vbObjectError + 701, "McInstallAll", why
    For Each nm In mBundleOrder
        all.Add mBundle(nm)
    Next nm
    If Failed Is Nothing Then Set Failed = New Collection
    If wrapped Is Nothing Then Set wrapped = New Collection
    McInstallAll = modLambdaLib.AddEntriesToWorkbook(all, wb, Failed, wrapped)
    If Failed.count = 0 Then StampVersion wb
End Function

' The library functions a workbook's formulas actually call - on any sheet, or
' inside a defined name of the user's own. Feed the result to
' McEnsureFunctions to install or repair exactly what the workbook needs.
Public Function McNamesUsedIn(ByVal wb As Workbook) As Collection
    Dim out As New Collection, seen As Object
    Dim ws As Worksheet, rg As Range, area As Range, v As Variant, item As Variant
    Dim nm As name

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    For Each ws In wb.Worksheets
        Set rg = Nothing
        On Error Resume Next
        Set rg = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not rg Is Nothing Then
            For Each area In rg.Areas
                v = area.Formula2
                If IsArray(v) Then
                    For Each item In v
                        CollectNames CStr(item), seen, out
                    Next item
                Else
                    CollectNames CStr(v), seen, out
                End If
            Next area
        End If
    Next ws
    For Each nm In wb.names
        If StrComp(Left$(nm.name, 7), "fx.Risk", vbTextCompare) <> 0 Then
            CollectNames nm.RefersTo, seen, out
        End If
    Next nm
    Set McNamesUsedIn = out
End Function

Private Sub CollectNames(ByVal formula As String, ByVal seen As Object, ByVal out As Collection)
    Dim t As Variant
    If InStr(1, formula, "fx.Risk", vbTextCompare) = 0 Then Exit Sub
    For Each t In ReferencedNames(formula)
        If Not seen.Exists(CStr(t)) Then
            seen(CStr(t)) = True
            out.Add CStr(t)
        End If
    Next t
End Sub


' --- developer build steps (no dialogs here; modMonteCarloRibbon wraps them) --

' Load MonteCarlo.txt into the bundle sheet of THIS workbook - the add-in being
' built. Refuses a file whose fx.RiskVersion is not MC_LIBRARY_VERSION, so a
' bundle can never drift from the VBA it ships with. Does not save.
Public Function McLoadBundleFromFile(ByVal path As String, ByRef count As Long, _
                                     ByRef why As String) As Boolean
    Dim entries As Collection, e As Variant, ws As Worksheet
    Dim data() As Variant, i As Long, ver As String, wasAddin As Boolean

    count = 0
    why = ""
    If Len(Dir$(path)) = 0 Then
        why = "Cannot find " & path
        Exit Function
    End If
    Set entries = modLambdaLib.EntriesFromTextFile(path)
    If entries Is Nothing Then
        why = "Nothing could be read from " & path
        Exit Function
    End If
    If entries.count = 0 Then
        why = "Nothing could be read from " & path
        Exit Function
    End If

    For Each e In entries
        If StrComp(CStr(e(0)), McFunctionName("fx.RiskVersion"), vbTextCompare) = 0 Then
            ver = VersionFromFormula(CStr(e(1)))
        End If
    Next e
    If ver <> MC_LIBRARY_VERSION Then
        why = path & " is version " & IIf(Len(ver) > 0, ver, "(unknown)") & _
              ", but MC_LIBRARY_VERSION in modMonteCarlo is " & MC_LIBRARY_VERSION & "." & _
              " Update the constant (and check the chart block sizes) before bundling."
        Exit Function
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(BUNDLE_SHEET)
    On Error GoTo Failed
    If ws Is Nothing Then
        ' An .xlam will not take a new sheet while it is an add-in.
        wasAddin = ThisWorkbook.IsAddin
        If wasAddin Then ThisWorkbook.IsAddin = False
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = BUNDLE_SHEET
        ws.Visible = xlSheetVeryHidden
        If wasAddin Then ThisWorkbook.IsAddin = True
    End If

    ReDim data(1 To entries.count, 1 To 3)
    i = 0
    For Each e In entries
        i = i + 1
        data(i, 1) = e(0)
        data(i, 2) = e(1)
        data(i, 3) = e(2)
    Next e

    ws.Cells.Clear
    ws.Range("A1:C1").value = Array("Function", "Formula", "Description")
    ws.Range("E1").value = "Version"
    ws.Range(BUNDLE_VERSION_CELL).NumberFormat = "@"
    ws.Range(BUNDLE_VERSION_CELL).value = MC_LIBRARY_VERSION
    ' Text format FIRST, or Excel tries to evaluate "=LAMBDA(...)" on the way in.
    ws.Range("B2").Resize(entries.count, 1).NumberFormat = "@"
    ws.Range("C2").Resize(entries.count, 1).NumberFormat = "@"
    ws.Range("A2").Resize(entries.count, 3).value = data

    Set mBundle = Nothing                            ' next read comes off the sheet
    count = entries.count
    McLoadBundleFromFile = True
    Exit Function

Failed:
    If wasAddin Then ThisWorkbook.IsAddin = True
    why = "Error " & Err.Number & " writing the bundle: " & Err.description
End Function

' "0.5.0" out of =LAMBDA("0.5.0 (39 distributions, ...)")
Private Function VersionFromFormula(ByVal f As String) As String
    Dim p As Long, q As Long
    p = InStr(1, f, """")
    If p = 0 Then Exit Function
    q = p + 1
    Do While q <= Len(f)
        If Not (Mid$(f, q, 1) Like "[0-9.]") Then Exit Do
        q = q + 1
    Loop
    VersionFromFormula = Mid$(f, p + 1, q - p - 1)
End Function

' Remove the library's functions from XL Edge's own LAMBDA library
' (tblAlonzoChurch), where earlier builds kept them. Through AddInStorage's own
' writer, so the neighbouring tables are never at risk. Does not save.
Public Function McRemoveFromLambdaLibrary(ByRef removed As Long, ByRef why As String) As Boolean
    Dim all As Collection, keep As New Collection, e As Variant, nm As String
    removed = 0
    why = ""
    Set all = AddInStorage.LambdaAll()
    For Each e In all
        nm = CStr(e(0))
        If StrComp(Left$(nm, 7), "fx.Risk", vbTextCompare) = 0 And Right$(nm, 1) = Lam() Then
            removed = removed + 1
        Else
            keep.Add e
        End If
    Next e
    If removed = 0 Then
        McRemoveFromLambdaLibrary = True
        Exit Function
    End If
    If AddInStorage.WriteLambdaBlock(keep) Then
        McRemoveFromLambdaLibrary = True
    Else
        why = AddInStorage.LastError
    End If
End Function


' --- workbook-level settings -------------------------------------------------

' MC_Trials drives every distribution's trial count, so a model moves from
' 1,000 to 100,000 trials by editing one thing. It is created as a constant
' name; point it at a cell instead if you would rather edit it on the grid.
Public Function McTrialsDefined(ByVal wb As Workbook) As Boolean
    Dim s As String
    On Error Resume Next
    s = wb.names(MC_TRIALS_NAME).RefersTo
    On Error GoTo 0
    McTrialsDefined = (Len(s) > 0)
End Function

' The trial count MC_Trials holds, or 0 when it is missing or points somewhere
' this cannot read. Used only to say how far a result will spill.
Public Function McTrialsCount(ByVal wb As Workbook) As Long
    McTrialsCount = CLng(NameValue(wb, MC_TRIALS_NAME, 0))
End Function

Public Function McDefineTrials(ByVal wb As Workbook, ByVal trials As Long) As String
    SetConstantName wb, MC_TRIALS_NAME, CStr(trials)
    McDefineTrials = MC_TRIALS_NAME
End Function

' Hands out the next unused stream id and advances the counter. Two inputs
' sharing a VarID would draw the IDENTICAL uniform stream and move together
' perfectly, which is almost never what anyone wants - so this is not optional
' bookkeeping, it is what keeps the inputs independent.
Public Function McNextVarId(ByVal wb As Workbook) As Long
    Dim nxt As Long
    nxt = CLng(NameValue(wb, MC_VARID_NAME, 1))
    If nxt < 1 Then nxt = 1
    SetConstantName wb, MC_VARID_NAME, CStr(nxt + 1)
    McNextVarId = nxt
End Function

Private Function NameValue(ByVal wb As Workbook, ByVal nm As String, _
                           ByVal fallback As Double) As Double
    Dim s As String
    NameValue = fallback
    On Error Resume Next
    s = wb.names(nm).RefersTo
    On Error GoTo 0
    If Len(s) < 2 Then Exit Function
    If Left$(s, 1) = "=" Then s = Mid$(s, 2)
    If IsNumeric(s) Then NameValue = CDbl(s)
End Function

Private Sub SetConstantName(ByVal wb As Workbook, ByVal nm As String, ByVal valueText As String)
    ' Delete rather than blank: assigning "" to RefersTo raises 1004.
    On Error Resume Next
    wb.names(nm).Delete
    On Error GoTo 0
    wb.names.Add name:=nm, RefersTo:="=" & valueText
End Sub


' --- building and writing the formula ---------------------------------------

' Application.InputBox with Type:=0 hands back a formula: "100" when the user
' types a number, "=$B$4" when they click a cell. Strip the leading sign so the
' fragment can be dropped into a larger formula either way.
Public Function McArgText(ByVal raw As String) As String
    Dim s As String
    s = Trim$(raw)
    If Left$(s, 1) = "=" Then s = Mid$(s, 2)
    McArgText = Trim$(s)
End Function

' A catalog parameter beginning "?" is optional. These split that marker off.
Public Function McParamIsOptional(ByVal param As String) As Boolean
    McParamIsOptional = (Left$(param, 1) = "?")
End Function

Public Function McParamLabel(ByVal param As String) As String
    If McParamIsTrailing(param) Then
        McParamLabel = Mid$(param, 3)
    ElseIf McParamIsOptional(param) Then
        McParamLabel = Mid$(param, 2)
    Else
        McParamLabel = param
    End If
End Function

' "?^..." - an optional parameter that comes AFTER Trials, VarID and Seed.
Public Function McParamIsTrailing(ByVal param As String) As Boolean
    McParamIsTrailing = (Left$(param, 2) = "?^")
End Function

' True when the user typed "none" for an optional parameter. InputBox Type:=0
' may hand it back as none, =none or ="none", so all three are accepted.
Public Function McIsNone(ByVal argText As String) As Boolean
    McIsNone = (LCase$(Replace(McArgText(argText), """", "")) = "none")
End Function

' A labeled VaR / CVaR / ES block for a spilled range of trials:
'
'   =LET(tr, A2#, cl, 0.95, VSTACK(HSTACK("Confidence", cl),
'        HSTACK("VaR", fx.RiskVaR<lambda>(tr, cl)), ... ))
'
' LET names avoid anything that reads as a cell reference; "c" alone would be
' refused, because Excel takes it as R1C1 notation for a column.
Public Function McBuildRiskBlock(ByVal ref As String, ByVal confidence As Double, _
                                 ByVal lossesPositive As Boolean, ByVal nameRef As String) As String
    Dim conf As String

    conf = Trim$(Str$(confidence))          ' Str$ always uses a "." decimal point
    If Left$(conf, 1) = "." Then conf = "0" & conf

    ' =fx.RiskMeasures<lambda>($D$2#, 0.95, , $D$1) - an empty argument leaves
    ' LossesPositive at its default, P&L.
    McBuildRiskBlock = "=" & McFunctionName("fx.RiskMeasures") & "(" & ref & ", " & conf & _
                       ", " & IIf(lossesPositive, "TRUE", "") & ", " & nameRef & ")"
End Function

' The last nTrailing args go after the Trials / VarID / Seed tail, with Seed
' left empty:  =fx.RiskCopulaClayton<lambda>(2, , MC_Trials, 31, , 180)
' Trailing args that were all left out are not written at all.
Public Function McBuildCall(ByVal baseName As String, ByVal args As Variant, _
                            ByVal trialsExpr As String, ByVal varId As Long, _
                            Optional ByVal nTrailing As Long = 0) As String
    Dim s As String, tail As String, i As Long, lastLead As Long

    lastLead = UBound(args) - nTrailing
    s = "=" & McFunctionName(baseName) & "("
    For i = LBound(args) To lastLead
        If i > LBound(args) Then s = s & ", "
        s = s & CStr(args(i))
    Next i
    If Len(trialsExpr) > 0 Then s = s & ", " & trialsExpr & ", " & CStr(varId)
    For i = lastLead + 1 To UBound(args)
        tail = tail & ", " & CStr(args(i))
    Next i
    If Len(Replace(Replace(tail, ",", ""), " ", "")) > 0 Then s = s & ", " & tail
    McBuildCall = s & ")"
End Function

' A reference to the whole spilled result - "'Model'!$A$2#" - given any cell of
' it. Sheet-qualified only when the target lives elsewhere, because an
' unqualified reference reads better on the same sheet.
Public Function McSpillReference(ByVal source As Range, ByVal target As Range) As String
    Dim anchor As Range, ref As String

    Set anchor = SpillAnchor(source)
    If anchor Is Nothing Then Exit Function

    ref = anchor.Address
    If anchor.Worksheet.name <> target.Worksheet.name Then
        ref = "'" & anchor.Worksheet.name & "'!" & ref
    End If
    McSpillReference = ref & "#"
End Function

' The cell a spilled result starts from, given any cell of it.
Public Function McSpillAnchor(ByVal rg As Range) As Range
    Set McSpillAnchor = SpillAnchor(rg)
End Function

' What to call a result in a chart title: the text in the cell above its
' anchor ("Profit"), or failing that the anchor's address.
Public Function McResultName(ByVal anchor As Range) As String
    Dim above As Variant
    If anchor.Row > 1 Then
        above = anchor.offset(-1, 0).Value2
        If VarType(above) = vbString Then
            If Len(Trim$(above)) > 0 Then
                McResultName = Trim$(above)
                Exit Function
            End If
        End If
    End If
    McResultName = anchor.Address(False, False)
End Function

' A formula fragment naming an input for fx.RiskTornado: a LIVE reference to
' the header cell above its anchor when there is one, so renaming the header
' renames the bar; otherwise the anchor's address as quoted text.
Public Function McLabelFragment(ByVal anchor As Range, ByVal target As Range) As String
    Dim above As Variant
    If anchor.Row > 1 Then
        above = anchor.offset(-1, 0).Value2
        If VarType(above) = vbString Then
            If Len(Trim$(above)) > 0 Then
                McLabelFragment = McQualifiedAddress(anchor.offset(-1, 0), target)
                Exit Function
            End If
        End If
    End If
    McLabelFragment = """" & anchor.Address(False, False) & """"
End Function

' "$A$1", or "'Model'!$A$1" when it lives on a different sheet from target.
Public Function McQualifiedAddress(ByVal cell As Range, ByVal target As Range) As String
    McQualifiedAddress = cell.Address
    If cell.Worksheet.name <> target.Worksheet.name Then
        McQualifiedAddress = "'" & cell.Worksheet.name & "'!" & McQualifiedAddress
    End If
End Function

' =fx.RiskTornado<lambda>(out#, HSTACK(a#, b#), HSTACK($A$1, $B$1), "rank", , $H$1)
' outLabel names the output for the chart title: a header-cell reference or a
' quoted address, from McLabelFragment.
Public Function McBuildTornado(ByVal outRef As String, ByVal inRefs As Collection, _
                               ByVal labels As Collection, ByVal useRank As Boolean, _
                               ByVal outLabel As String) As String
    McBuildTornado = "=" & McFunctionName("fx.RiskTornado") & "(" & outRef & _
                     ", HSTACK(" & JoinCollection(inRefs) & ")" & _
                     ", HSTACK(" & JoinCollection(labels) & ")" & _
                     IIf(useRank, ", ""rank""", ", ") & ", , " & outLabel & ")"
End Function

' =fx.RiskVariablesTable<lambda>(HSTACK(a#, b#), , HSTACK($A$1, $B$1))
' baseName is fx.RiskVariablesTable or fx.RiskVariablesTableDetail. Each name
' is a live header cell where there is one, as for the tornado's inputs.
Public Function McBuildVariablesTable(ByVal baseName As String, ByVal refs As Collection, _
                                      ByVal labels As Collection) As String
    McBuildVariablesTable = "=" & McFunctionName(baseName) & _
                            "(HSTACK(" & JoinCollection(refs) & ")" & _
                            ", , HSTACK(" & JoinCollection(labels) & "))"
End Function

Private Function JoinCollection(ByVal c As Collection) As String
    Dim v As Variant, s As String
    For Each v In c
        If Len(s) > 0 Then s = s & ", "
        s = s & CStr(v)
    Next v
    JoinCollection = s
End Function

Private Function SpillAnchor(ByVal rg As Range) As Range
    Dim c As Range
    If rg Is Nothing Then Exit Function
    Set c = rg.Cells(1, 1)

    ' SpillParent raises when the cell is not part of a spilled result, in which
    ' case the cell is its own anchor.
    On Error Resume Next
    Set SpillAnchor = c.SpillParent
    On Error GoTo 0
    If SpillAnchor Is Nothing Then Set SpillAnchor = c
End Function

' Formula2 rather than Formula. Writing a dynamic-array formula through the old
' Formula property makes Excel insert implicit-intersection @ signs, which turn
' a spilling call into a single value - the exact problem that stopped
' ProbabilityManagement wiring their own LAMBDAs up. Fall back only if this
' build of Excel has no Formula2 at all.
Public Sub McWriteFormula(ByVal target As Range, ByVal formulaText As String)
    Dim cell As Range
    Set cell = target.Cells(1, 1)

    On Error Resume Next
    cell.Formula2 = formulaText
    If Err.Number <> 0 Then
        Err.Clear
        cell.formula = formulaText
    End If
    On Error GoTo 0
End Sub
