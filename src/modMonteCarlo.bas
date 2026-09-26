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
' WHAT THIS FEATURE ACTUALLY DOES. It writes a native formula into a cell:
'
'   =fx.RiskPertλ($C$4, $C$5, $C$6, MC_Trials, 7)
'
' and nothing more. The formula calls LAMBDA functions stored as defined names
' in the user's own workbook, so the finished file keeps calculating for people
' who have never heard of XL Edge. The add-in is an AUTHORING tool. If it were
' ever uninstalled, every model built with it would carry on working.
'
' That is the whole point of the exercise - the add-ins this replaces all fail
' with #NAME? the moment the recipient does not have them installed.
'
' THE LAMBDA CHARACTER. Every function name ends in λ (U+03BB). It is NOT
' written literally anywhere in this module. A .bas file exports as Windows-1252
' and λ has no representation in that code page, so a literal would be silently
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

' Presence of this name is how we tell whether the library is installed.
Private Const PROBE_BASE As String = "fx.RiskU"

Private mCatalog As Collection


' --- the lambda suffix -------------------------------------------------------

' U+03BB GREEK SMALL LETTER LAMBDA. See the note in the module header for why
' this is ChrW and not a literal.
Public Function Lam() As String
    Lam = ChrW$(&H3BB)
End Function

Public Function McFunctionName(ByVal baseName As String) As String
    McFunctionName = baseName & Lam()
End Function


' --- the distribution catalogue ---------------------------------------------

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
    McCatalogCount = McCatalog().Count
End Function

Public Function McCatalogItem(ByVal idx As Long) As Variant
    If idx < 1 Or idx > McCatalogCount() Then Exit Function
    McCatalogItem = McCatalog().Item(idx)
End Function

' Maps a ribbon gallery item id ("McD_Pert") to its catalogue index.
' Returns 0 when the id is not in the catalogue.
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


' --- the library itself ------------------------------------------------------

Public Function McLibraryInstalled(ByVal wb As Workbook) As Boolean
    Dim s As String
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    s = wb.names(McFunctionName(PROBE_BASE)).RefersTo
    On Error GoTo 0
    McLibraryInstalled = (Len(s) > 0)
End Function

' Downloads the published module and installs every definition into wb.
' Raises on a failed download - the ribbon layer turns that into a message.
'
' NOT called McInstallLibrary: that is the ribbon callback's name. The ribbon
' finds callbacks by bare name across the whole project, and two Public
' procedures sharing a name make it give up with "Cannot run the macro".
Public Function McDownloadLibrary(ByVal wb As Workbook, _
                                 Optional ByRef failed As Collection, _
                                 Optional ByRef wrapped As Collection) As Long
    Dim entries As Collection

    If wb Is Nothing Then Exit Function
    Set entries = modLambdaLib.EntriesFromGistUrl(MC_GIST_URL)
    If entries Is Nothing Then Exit Function
    If entries.Count = 0 Then
        Err.Raise vbObjectError + 700, "McDownloadLibrary", _
                  "The published library downloaded but contained no definitions."
    End If

    If failed Is Nothing Then Set failed = New Collection
    If wrapped Is Nothing Then Set wrapped = New Collection
    McDownloadLibrary = modLambdaLib.AddEntriesToWorkbook(entries, wb, failed, wrapped)
End Function

' Installs only when the library is not already there. Returns True if, by the
' time it finishes, the workbook can evaluate the functions.
Public Function McEnsureLibrary(ByVal wb As Workbook, ByRef added As Long) As Boolean
    added = 0
    If McLibraryInstalled(wb) Then
        McEnsureLibrary = True
        Exit Function
    End If
    added = McDownloadLibrary(wb)
    McEnsureLibrary = McLibraryInstalled(wb)
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

' A catalogue parameter beginning "?" is optional. These split that marker off.
Public Function McParamIsOptional(ByVal param As String) As Boolean
    McParamIsOptional = (Left$(param, 1) = "?")
End Function

Public Function McParamLabel(ByVal param As String) As String
    If McParamIsOptional(param) Then McParamLabel = Mid$(param, 2) Else McParamLabel = param
End Function

' True when the user typed "none" for an optional parameter. InputBox Type:=0
' may hand it back as none, =none or ="none", so all three are accepted.
Public Function McIsNone(ByVal argText As String) As Boolean
    McIsNone = (LCase$(Replace(McArgText(argText), """", "")) = "none")
End Function

' A labelled VaR / CVaR / ES block for a spilled range of trials:
'
'   =LET(tr, A2#, cl, 0.95, VSTACK(HSTACK("Confidence", cl),
'        HSTACK("VaR", fx.RiskVaR<lambda>(tr, cl)), ... ))
'
' LET names avoid anything that reads as a cell reference; "c" alone would be
' refused, because Excel takes it as R1C1 notation for a column.
Public Function McBuildRiskBlock(ByVal ref As String, ByVal confidence As Double, _
                                 ByVal lossesPositive As Boolean) As String
    Dim tail As String, conf As String

    conf = Trim$(Str$(confidence))          ' Str$ always uses a "." decimal point
    If Left$(conf, 1) = "." Then conf = "0" & conf
    tail = IIf(lossesPositive, ", TRUE", "")

    McBuildRiskBlock = "=LET(tr, " & ref & ", cl, " & conf & ", VSTACK(" & _
        "HSTACK(""Confidence"", cl), " & _
        "HSTACK(""VaR"", " & McFunctionName("fx.RiskVaR") & "(tr, cl" & tail & ")), " & _
        "HSTACK(""CVaR"", " & McFunctionName("fx.RiskCVaR") & "(tr, cl" & tail & ")), " & _
        "HSTACK(""ES"", " & McFunctionName("fx.RiskES") & "(tr, cl" & tail & "))))"
End Function

Public Function McBuildCall(ByVal baseName As String, ByVal args As Variant, _
                            ByVal trialsExpr As String, ByVal varId As Long) As String
    Dim s As String, i As Long

    s = "=" & McFunctionName(baseName) & "("
    For i = LBound(args) To UBound(args)
        If i > LBound(args) Then s = s & ", "
        s = s & CStr(args(i))
    Next i
    If Len(trialsExpr) > 0 Then s = s & ", " & trialsExpr & ", " & CStr(varId)
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
