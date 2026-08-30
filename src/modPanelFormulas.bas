Attribute VB_Name = "modPanelFormulas"
Option Explicit

' ============================================================================
'  modPanelFormulas
'  Where every cell of a panel chart lives, and what formula goes in it.
'
'  Pure string work - no Excel, no COM, no worksheet. Like modPanelSpec this
'  module runs with every workbook closed, which is what lets the formulas be
'  diffed against the six shipped reference builds without opening Excel.
'
'  TWO ADDRESS MODES, ONE COMPOSER
'  -------------------------------
'  The Python engine this ports from writes A1 formulas, one distinct string
'  per row, because it bakes each row number into each formula. VBA does not
'  have to: every row of the plot table has the same formula SHAPE, so in R1C1
'  it is literally the same string and a whole column goes out in ONE
'  assignment. A 5x4x12x2 grid is ~250 COM calls that way against ~15,000 cell
'  by cell, which is the difference between instant and abandoned.
'
'  The cost is that every reference becomes an OFFSET, so being one column out
'  reads as a plausible chart rather than as an error. That is why both modes
'  exist and why they share one composer: PanelSetAddressMode(amA1) makes this
'  module emit the exact A1 strings quoted in
'
'      skills\panel-charts\references\examples\*.md
'
'  at the exact addresses printed beside them. Build a small grid each way,
'  diff the two plot tables cell for cell, and a disagreement is an offset bug
'  with its location attached.
'
'  WHAT MUST NOT CHANGE
'  --------------------
'  * A gap is NA(), never "". An empty string plots as a ZERO in both line and
'    scatter series, dropping a spike to the baseline.
'  * The INDEX guards are NESTED IFs, never OR(). OR() evaluates every one of
'    its arguments, so on a spacer slot the inner INDEX reaches outside the
'    input block and every gap cell shows #REF!.
'  * Stem columns write 0, not NA(), where the series does not own the point.
'    An error-bar range must be entirely numeric, the same rule that governs
'    XValues, and breaking it makes Excel discard the range without saying so.
'  * Every embedded number goes through NumStr, never CStr. On a comma-decimal
'    locale CStr(0.82) is "0,82", which then breaks a US-English formula. The
'    Python never had this problem because Python's %s on a float is always
'    dot-decimal.
'
'  Ported from panel_excel.py (PanelBuilder.__init__ and the _write_* formula
'  composers) in the panel-charts skill. Keep the two in step.
' ============================================================================


' ----------------------------------------------------------------------------
'  Style
' ----------------------------------------------------------------------------

Public Type TPanelStyle
    elementColours()  As Long       ' RGB longs, cycled; unallocated = palette
    signColourUp      As Long       ' pin, value >= 0
    signColourDown    As Long       ' pin, value < 0
    signColoursSet    As Boolean
    markerSize        As Long
    markerStyle       As Long       ' 0 = the kind's own head
    valueLabels       As Boolean
    chartTitle        As Boolean
    valueTicks        As Boolean
    labelFormat       As String
    labelSize         As Long
    titleSize         As Long
    dividerColour     As Long
    dividerWeight     As Double
    dividerAfter      As Long       ' 0 = none; a 1-based period index
    stemWeight        As Double
End Type

' ----------------------------------------------------------------------------
'  DECLARATIONS SECTION
'
'  EVERYTHING module-level lives here, above the first procedure: Type, Enum,
'  Const and module variables alike. VBA has no equivalent of a declaration
'  further down the file - the Declarations section ends at the first Sub or
'  Function, and a Type placed after one simply never registers.
'
'  That failure is invisible at the declaration. The error surfaces as
'  "User-defined type not defined" against the first procedure that NAMES the
'  type - LayoutInit, three hundred lines away - so it reads as a problem with
'  the consumer rather than with where the Type was put.
' ----------------------------------------------------------------------------

' ----------------------------------------------------------------------------
'  Parameter rows
'
'  Indices into the engine's parameter block, 0-based from its first row. The
'  block is exactly PARAM_ROWS tall and the plot table starts below it; an
'  earlier version of the Python guessed the gap, the plot-table header landed
'  on `span`, and every value formula then divided by text. LayoutInit asserts
'  the count rather than trusting it.
' ----------------------------------------------------------------------------

Public Const PARAM_ROWS As Long = 18

Public Const PR_GRIDROWS As Long = 0
Public Const PR_GRIDCOLS As Long = 1
Public Const PR_PERIODS As Long = 2
Public Const PR_ELEMENTS As Long = 3
Public Const PR_BLOCKS As Long = 4
Public Const PR_BANDS As Long = 5
Public Const PR_SUBSLOTS As Long = 6
Public Const PR_BLOCKSLOTS As Long = 7
Public Const PR_SPACERS As Long = 8
Public Const PR_LEAD As Long = 9
Public Const PR_BANDFRAC As Long = 10
Public Const PR_MINRAW As Long = 11
Public Const PR_MAXRAW As Long = 12
Public Const PR_STEP As Long = 13
Public Const PR_VMIN As Long = 14
Public Const PR_VMAX As Long = 15
Public Const PR_SPAN As Long = 16
Public Const PR_ZOFF As Long = 17

' The pad above vmax, as a share of the data's own range. Panel titles sit in
' it. Not a tunable - a title needs somewhere to go on every grid this builds.
Private Const TITLE_HEADROOM As Double = 0.1

' Layout invariants. See the header note on PARAM_ROWS: the test for a
' settings-sheet constant is "would a user ever want a different value, and
' does the code still work if they pick one?" - here, no and no.
Private Const IN_R0 As Long = 6
Private Const IN_C0 As Long = 2
Private Const HDR_ELEM_ROW As Long = 4
Private Const HDR_PER_ROW As Long = 5
Private Const PAR_R0 As Long = 3
Private Const ENG_GAP As Long = 3


' ----------------------------------------------------------------------------
'  Layout - every address the engine writes to
' ----------------------------------------------------------------------------

Public Type TPanelLayout
    originShift As Long          ' rows the whole engine is pushed down by

    ' input block
    inR0 As Long
    inC0 As Long
    inR1 As Long
    inC1 As Long
    inW As Long
    hdrElemRow As Long
    hdrPerRow As Long

    ' engine
    engC0 As Long                ' label column
    parC0 As Long                ' value column, engC0 + 1
    parR0 As Long

    ' plot table
    ptR0 As Long
    ptR1 As Long
    ptC0 As Long
    cX As Long
    cPos As Long
    cBlk As Long
    cSub As Long
    cK As Long
    cE As Long
    cSep As Long
    cKx As Long
    cBlkx As Long
    cCat As Long
    vc0 As Long                  ' first band column

    tones As Long                ' 2 for pin, else 1
    perBand As Long              ' columns one band consumes
    cMask As Long                ' area only; 0 = none

    ' stacked totals, under the input block they summarise
    ssR0 As Long
    ssR1 As Long
    ssC0 As Long

    errC0 As Long                ' pin stems; 0 = none
    labC0 As Long                ' point labels; 0 = none
    anC0 As Long                 ' first annotation column

    chartRow As Long
    labelEvery As Long
End Type
' ============================================================================
'  Address helpers
' ============================================================================

Public Enum eAddrMode
    amR1C1 = 0
    amA1 = 1
End Enum

Private mMode As eAddrMode

' The engine's own defaults are notation-neutral: colour by element ordinal, no
' markers, no value labels. That is right for a chart whose series are just
' series. A caller drawing in a notation - IBCS, say - overrides them, because
' there a fill is not decoration but a claim about what kind of number this is.
Public Function PanelStyleDefaults() As TPanelStyle
    Dim st As TPanelStyle
    st.markerSize = 5
    st.markerStyle = 0
    st.valueLabels = False
    st.chartTitle = True
    st.valueTicks = True
    st.labelFormat = "#,##0"
    st.labelSize = 8
    st.titleSize = 9
    st.dividerColour = RGB(191, 191, 191)
    st.dividerWeight = 0.75
    st.dividerAfter = 0
    st.stemWeight = 1.5
    st.signColourUp = RGB(140, 180, 0)
    st.signColourDown = RGB(192, 0, 0)
    st.signColoursSet = False
    PanelStyleDefaults = st
End Function

' The house palette, used when the caller supplies no element colours.
Public Function PanelPaletteColour(ByVal element As Long) As Long
    Dim n As Long
    n = element Mod 8
    Select Case n
        Case 0: PanelPaletteColour = RGB(31, 78, 121)
        Case 1: PanelPaletteColour = RGB(128, 128, 128)
        Case 2: PanelPaletteColour = RGB(46, 117, 182)
        Case 3: PanelPaletteColour = RGB(237, 125, 49)
        Case 4: PanelPaletteColour = RGB(112, 173, 71)
        Case 5: PanelPaletteColour = RGB(191, 143, 0)
        Case 6: PanelPaletteColour = RGB(165, 165, 165)
        Case Else: PanelPaletteColour = RGB(68, 114, 196)
    End Select
End Function

Public Function PanelElementColour(ByRef st As TPanelStyle, _
                                   ByVal element As Long) As Long
    Dim n As Long
    On Error GoTo UsePalette
    n = UBound(st.elementColours) - LBound(st.elementColours) + 1
    If n < 1 Then GoTo UsePalette
    PanelElementColour = st.elementColours(LBound(st.elementColours) + (element Mod n))
    Exit Function
UsePalette:
    PanelElementColour = PanelPaletteColour(element)
End Function


' Work out where everything goes. Pure arithmetic; call it before anything
' reads an address. Returns "" or a problem.
Public Function LayoutInit(ByRef sp As TPanelSpec, ByRef ly As TPanelLayout, _
                           Optional ByVal originShift As Long = 0) As String

    Dim o As Long
    o = originShift
    ly.originShift = o

    ' --- input block
    ly.inR0 = IN_R0 + o
    ly.inC0 = IN_C0
    ly.inW = sp.Elements * sp.Periods
    ly.inR1 = ly.inR0 + PS_NPanels(sp) - 1
    ly.inC1 = ly.inC0 + ly.inW - 1
    ly.hdrElemRow = HDR_ELEM_ROW + o
    ly.hdrPerRow = HDR_PER_ROW + o

    ' --- engine. Three clear columns so a wide input block never collides
    '     with the parameters.
    ly.engC0 = ly.inC1 + ENG_GAP
    ly.parC0 = ly.engC0 + 1
    ly.parR0 = PAR_R0 + o

    ' The plot table shares the engine's columns, so it must start clear of the
    ' parameter block - hence PARAM_ROWS rather than a guessed gap, plus a
    ' header row and some breathing room.
    ly.ptR0 = ly.parR0 + PARAM_ROWS + 3
    ly.ptC0 = ly.engC0
    ly.ptR1 = ly.ptR0 + PS_NSlots(sp) - 1

    ' --- helper columns, in the order the Python lays them out
    ly.cX = ly.ptC0
    ly.cPos = ly.ptC0 + 1
    ly.cBlk = ly.ptC0 + 2
    ly.cSub = ly.ptC0 + 3
    ly.cK = ly.ptC0 + 4
    ly.cE = ly.ptC0 + 5
    ly.cSep = ly.ptC0 + 6
    ' Effective period and block: what a slot should LOOK UP, as opposed to
    ' where it sits. They differ for bar, whose periods are written in reverse
    ' so period 1 reads at the top, and for area, whose spacer and lead slots
    ' hold a neighbouring value flat instead of going absent.
    ly.cKx = ly.ptC0 + 7
    ly.cBlkx = ly.ptC0 + 8
    ly.cCat = ly.ptC0 + 9
    ly.vc0 = ly.ptC0 + 10

    ' A pin needs two value columns per element - one per direction - and a
    ' stem-length column for each, because an error bar's length is data and
    ' has to live in cells like everything else.
    If PS_TwoTone(sp) Then ly.tones = 2 Else ly.tones = 1
    ly.perBand = sp.Elements * ly.tones
    If PS_Stacked(sp) Then ly.perBand = ly.perBand + 1

    Dim afterValues As Long
    afterValues = ly.vc0 + PS_Bands(sp) * ly.perBand

    If PS_StacksElements(sp) Then
        ly.cMask = afterValues
        afterValues = afterValues + 2
    Else
        ly.cMask = 0
        afterValues = afterValues + 1
    End If

    ' Stacked totals sit under the input block they summarise, clear of the
    ' engine columns - beside the parameters they collided with the plot table
    ' as soon as the grid grew past about twenty panels.
    ly.ssC0 = ly.inC0
    ly.ssR0 = ly.inR1 + 3
    If PS_StacksElements(sp) Then
        ly.ssR1 = ly.ssR0 + PS_NPanels(sp) - 1
    Else
        ly.ssR1 = ly.inR1
    End If

    ' Stem lengths and label text are separate blocks because not every kind
    ' that needs one needs the other: a pin has a stem AND a label, a dot has a
    ' label and no stem, and the rest have neither.
    Dim cursor As Long
    cursor = afterValues
    ly.errC0 = 0
    ly.labC0 = 0
    If PS_TwoTone(sp) Then
        ly.errC0 = cursor
        cursor = cursor + PS_Bands(sp) * ly.perBand + 1
    End If
    If PS_Labelled(sp) Then
        ly.labC0 = cursor
        cursor = cursor + PS_Bands(sp) * ly.perBand + 1
    End If
    ly.anC0 = cursor

    ' The chart is parked below every cell the engine writes, so the print area
    ' under it stays empty and nothing prints through it.
    ly.chartRow = ly.inR1
    If ly.ptR1 > ly.chartRow Then ly.chartRow = ly.ptR1
    If ly.ssR1 > ly.chartRow Then ly.chartRow = ly.ssR1
    ly.chartRow = ly.chartRow + 3

    ' Print every period label on a small grid; thin them out on a wide one,
    ' where the slots are too narrow to carry them all.
    If PS_Blocks(sp) * sp.Periods <= 24 Then
        ly.labelEvery = 1
    Else
        ly.labelEvery = sp.Periods \ 4
        If ly.labelEvery < 1 Then ly.labelEvery = 1
    End If

    LayoutInit = ""
End Function


' Column of the invisible base series for a band (stacked kinds only).
Public Function LY_BaseCol(ByRef ly As TPanelLayout, ByVal band As Long) As Long
    LY_BaseCol = ly.vc0 + band * ly.perBand
End Function

Public Function LY_ValCol(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                          ByVal band As Long, ByVal element As Long, _
                          Optional ByVal tone As Long = 0) As Long
    Dim off As Long
    If PS_Stacked(sp) Then off = 1 Else off = 0
    LY_ValCol = ly.vc0 + band * ly.perBand + off + element * ly.tones + tone
End Function

' Stem length for one pin series, in value-axis units.
Public Function LY_ErrCol(ByRef ly As TPanelLayout, ByVal band As Long, _
                          ByVal element As Long, _
                          Optional ByVal tone As Long = 0) As Long
    LY_ErrCol = ly.errC0 + band * ly.perBand + element * ly.tones + tone
End Function

' The formatted text each point's label is linked to.
Public Function LY_LabCol(ByRef ly As TPanelLayout, ByVal band As Long, _
                          ByVal element As Long, _
                          Optional ByVal tone As Long = 0) As Long
    LY_LabCol = ly.labC0 + band * ly.perBand + element * ly.tones + tone
End Function


' amR1C1 is the shipping mode. amA1 exists so the generated formulas can be
' compared, string for string, against the addresses printed in the reference
' documents - see the module header.
Public Sub PanelSetAddressMode(ByVal m As eAddrMode)
    mMode = m
End Sub

Public Function PanelAddressMode() As eAddrMode
    PanelAddressMode = mMode
End Function

' Excel's "" - two doubled quotes, which is four characters of VBA. Named,
' because scattering them through the composers is how a quote goes missing.
Public Property Get EMPTYTXT() As String
    EMPTYTXT = """"""
End Property

' Wrap a string as an Excel string literal, doubling any quote inside it.
Public Function Q(ByVal s As String) As String
    Q = Chr$(34) & Replace(s, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

' A number, in a formula, on any locale.
'
' CStr(0.82) on a comma-decimal machine is "0,82", which then goes into a
' US-English formula and breaks it. Format$ with an explicit picture is
' locale-independent. This is the single most likely source of "works here,
' error 1004 there", and the Python had no equivalent hazard.
Public Function NumStr(ByVal v As Double) As String
    NumStr = Format$(v, "0.############")
End Function

Public Function ColLetters(ByVal col As Long) As String
    Dim s As String, c As Long, r As Long
    c = col
    Do While c > 0
        r = (c - 1) Mod 26
        s = Chr$(65 + r) & s
        c = (c - 1) \ 26
    Loop
    ColLetters = s
End Function

' A same-row reference from the column being written to another column.
'
' In R1C1 this is an OFFSET, which is what makes a whole column one assignment
' and also what makes an off-by-one look like a plausible chart. Every caller
' passes both columns by name off TPanelLayout; a literal RC[-3] never appears
' in this module.
Public Function Ref(ByVal fromCol As Long, ByVal toCol As Long, _
                    ByVal r As Long) As String
    If mMode = amA1 Then
        Ref = ColLetters(toCol) & r
    ElseIf fromCol = toCol Then
        Ref = "RC"
    Else
        Ref = "RC[" & (toCol - fromCol) & "]"
    End If
End Function

' An absolute reference to one of the engine's parameter cells.
Public Function Par(ByRef ly As TPanelLayout, ByVal idx As Long) As String
    Dim r As Long
    r = ly.parR0 + idx
    If mMode = amA1 Then
        Par = "$" & ColLetters(ly.parC0) & "$" & r
    Else
        Par = "R" & r & "C" & ly.parC0
    End If
End Function

' An absolute rectangular range.
Public Function Rgn(ByVal r1 As Long, ByVal c1 As Long, _
                    ByVal r2 As Long, ByVal c2 As Long) As String
    If mMode = amA1 Then
        Rgn = "$" & ColLetters(c1) & "$" & r1 & ":$" & ColLetters(c2) & "$" & r2
    Else
        Rgn = "R" & r1 & "C" & c1 & ":R" & r2 & "C" & c2
    End If
End Function

' A same-row span of columns, used by the stacked base to sum a band.
Public Function RefSpan(ByVal fromCol As Long, ByVal c1 As Long, _
                        ByVal c2 As Long, ByVal r As Long) As String
    RefSpan = Ref(fromCol, c1, r) & ":" & Ref(fromCol, c2, r)
End Function

' The input block - the only cells the reader ever edits.
Public Function InRng(ByRef ly As TPanelLayout) As String
    InRng = Rgn(ly.inR0, ly.inC0, ly.inR1, ly.inC1)
End Function

Public Function PeriodRng(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec) As String
    PeriodRng = Rgn(ly.hdrPerRow, ly.inC0, ly.hdrPerRow, ly.inC0 + sp.Periods - 1)
End Function

Public Function StackSumRng(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec) As String
    StackSumRng = Rgn(ly.ssR0, ly.ssC0, ly.ssR1, ly.ssC0 + sp.Periods - 1)
End Function


' ============================================================================
'  The helper columns
'
'  Nine columns that turn a category slot number into "which panel, which
'  period, which element, and is this a gap". Everything downstream reads them
'  rather than recomputing the arithmetic, which is why the value columns stay
'  legible at all.
' ============================================================================

Public Function FxPos(ByRef ly As TPanelLayout, ByVal r As Long) As String
    FxPos = "=" & Ref(ly.cPos, ly.cX, r) & "-" & Par(ly, PR_LEAD)
End Function

Public Function FxBlk(ByRef ly As TPanelLayout, ByVal r As Long) As String
    Dim p As String, pitch As String
    p = Ref(ly.cBlk, ly.cPos, r)
    pitch = "(" & Par(ly, PR_BLOCKSLOTS) & "+" & Par(ly, PR_SPACERS) & ")"
    FxBlk = "=IF(" & p & "<=0,-1,INT((" & p & "-1)/" & pitch & "))"
End Function

Public Function FxSub(ByRef ly As TPanelLayout, ByVal r As Long) As String
    Dim p As String, pitch As String
    p = Ref(ly.cSub, ly.cPos, r)
    pitch = "(" & Par(ly, PR_BLOCKSLOTS) & "+" & Par(ly, PR_SPACERS) & ")"
    FxSub = "=IF(" & p & "<=0,-1,MOD(" & p & "-1," & pitch & "))"
End Function

Public Function FxK(ByRef ly As TPanelLayout, ByVal r As Long) As String
    Dim s As String
    s = Ref(ly.cK, ly.cSub, r)
    FxK = "=IF(OR(" & s & "<0," & s & ">=" & Par(ly, PR_BLOCKSLOTS) & ")," & _
          "-1,INT(" & s & "/" & Par(ly, PR_SUBSLOTS) & "))"
End Function

Public Function FxE(ByRef ly As TPanelLayout, ByVal r As Long) As String
    Dim s As String
    s = Ref(ly.cE, ly.cSub, r)
    FxE = "=IF(OR(" & s & "<0," & s & ">=" & Par(ly, PR_BLOCKSLOTS) & ")," & _
          "-1,MOD(" & s & "," & Par(ly, PR_SUBSLOTS) & "))"
End Function

' 1 on a slot that is not part of any panel: the lead margin, a spacer between
' blocks, or bar's trailing gap. Everything that draws checks this first.
Public Function FxSep(ByRef ly As TPanelLayout, ByVal r As Long) As String
    Dim p As String, s As String, b As String
    p = Ref(ly.cSep, ly.cPos, r)
    s = Ref(ly.cSep, ly.cSub, r)
    b = Ref(ly.cSep, ly.cBlk, r)
    FxSep = "=IF(OR(" & p & "<=0," & s & "<0," & s & ">=" & _
            Par(ly, PR_BLOCKSLOTS) & "," & b & ">=" & Par(ly, PR_BLOCKS) & _
            "),1,0)"
End Function

' Effective period - what this slot looks up.
Public Function FxKx(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                     ByVal r As Long) As String
    Dim p As String, s As String, k As String, sep As String
    p = Ref(ly.cKx, ly.cPos, r)
    s = Ref(ly.cKx, ly.cSub, r)
    k = Ref(ly.cKx, ly.cK, r)
    sep = Ref(ly.cKx, ly.cSep, r)

    If PS_StacksElements(sp) Then
        ' AREA. A stacked area cannot be gapped at all - Excel joins whatever
        ' points it is given - so instead of going absent the lead and spacer
        ' slots hold the NEIGHBOURING value flat, and a mask column paints over
        ' the join afterwards.
        FxKx = "=IF(" & p & "<=0,0,IF(" & s & "<" & Par(ly, PR_BLOCKSLOTS) & _
               "," & k & ",IF(" & s & "=" & Par(ly, PR_BLOCKSLOTS) & "," & _
               Par(ly, PR_PERIODS) & "-1,0)))"
    ElseIf PS_ReversePeriods(sp) Then
        ' BAR. Excel draws category 1 at the bottom, so period 1 is written to
        ' the LAST slot of its block and reads at the top.
        FxKx = "=IF(" & sep & "=1,-1," & Par(ly, PR_PERIODS) & "-1-" & k & ")"
    Else
        FxKx = "=IF(" & sep & "=1,-1," & k & ")"
    End If
End Function

' Effective block - which panel column this slot looks up.
Public Function FxBlkx(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                       ByVal r As Long) As String
    Dim p As String, s As String, b As String, sep As String
    p = Ref(ly.cBlkx, ly.cPos, r)
    s = Ref(ly.cBlkx, ly.cSub, r)
    b = Ref(ly.cBlkx, ly.cBlk, r)
    sep = Ref(ly.cBlkx, ly.cSep, r)

    If PS_StacksElements(sp) Then
        FxBlkx = "=IF(" & p & "<=0,0,IF(" & s & "<=" & _
                 Par(ly, PR_BLOCKSLOTS) & "," & b & "," & b & "+1))"
    Else
        FxBlkx = "=IF(" & sep & "=1,-1," & b & ")"
    End If
End Function

' The category axis label. Blank on a gap, blank on a period being skipped,
' and centred under its element cluster otherwise.
Public Function FxCat(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                      ByVal r As Long) As String
    Dim sep As String, kx As String, e As String
    sep = Ref(ly.cCat, ly.cSep, r)
    kx = Ref(ly.cCat, ly.cKx, r)
    e = Ref(ly.cCat, ly.cE, r)

    FxCat = "=IF(" & sep & "=1," & EMPTYTXT & ",IF(AND(MOD(" & kx & "," & _
            ly.labelEvery & ")=0," & e & "=INT((" & Par(ly, PR_SUBSLOTS) & _
            "-1)/2))," & "INDEX(" & PeriodRng(ly, sp) & "," & kx & "+1)," & _
            EMPTYTXT & "))"
End Function


' ============================================================================
'  The value columns
' ============================================================================

' The panel's 1-based row in the input block, mirroring PS_Place.
'
' Bar reverses the BLOCK and reads its band from the grid column; every other
' kind reverses the band and reads its block from the grid column. Getting this
' backwards fills every panel with a neighbour's numbers and raises nothing.
Public Function FxPanelExpr(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                            ByVal band As Long, ByVal blkRef As String) As String
    If PS_Transposed(sp) Then
        FxPanelExpr = "(" & Par(ly, PR_GRIDROWS) & "-1-" & blkRef & ")*" & _
                      Par(ly, PR_GRIDCOLS) & "+" & band & "+1"
    Else
        FxPanelExpr = "(" & Par(ly, PR_GRIDROWS) & "-1-" & band & ")*" & _
                      Par(ly, PR_GRIDCOLS) & "+" & blkRef & "+1"
    End If
End Function

' The raw input cell one slot of one series reads.
Public Function FxValueRef(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                           ByVal band As Long, ByVal element As Long, _
                           ByVal col As Long, ByVal r As Long) As String
    Dim blkx As String, kx As String
    blkx = Ref(col, ly.cBlkx, r)
    kx = Ref(col, ly.cKx, r)

    FxValueRef = "INDEX(" & InRng(ly) & "," & _
                 FxPanelExpr(ly, sp, band, blkx) & "," & _
                 element & "*" & Par(ly, PR_PERIODS) & "+" & kx & "+1)"
End Function

' Where a value lands on the hidden scaffolding axis.
'
' A stacked kind plots a HEIGHT above its band's zero line, because an
' invisible base series underneath carries the offset - hence "-0" rather than
' "-vmin". Everything else plots the absolute banded position.
Public Function FxPlaced(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                         ByVal band As Long, ByVal valueExpr As String) As String
    If PS_Stacked(sp) Then
        FxPlaced = "(" & valueExpr & "-0)/" & Par(ly, PR_SPAN) & "*" & _
                   Par(ly, PR_BANDFRAC)
    Else
        FxPlaced = band & "+(" & valueExpr & "-" & Par(ly, PR_VMIN) & ")/" & _
                   Par(ly, PR_SPAN) & "*" & Par(ly, PR_BANDFRAC)
    End If
End Function

' One data column: line, column, area and dot. Pin has its own composer.
Public Function FxValue(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                        ByVal band As Long, ByVal element As Long, _
                        ByVal r As Long) As String

    Dim col As Long
    col = LY_ValCol(ly, sp, band, element, 0)

    Dim value As String, placed As String
    Dim sep As String, e As String, blkx As String
    value = FxValueRef(ly, sp, band, element, col, r)
    placed = FxPlaced(ly, sp, band, value)
    sep = Ref(col, ly.cSep, r)
    e = Ref(col, ly.cE, r)
    blkx = Ref(col, ly.cBlkx, r)

    If PS_Labelled(sp) Then
        ' DOT. A grid may hold fewer panels than cells, and a blank input cell
        ' reads as a zero - a row of dots along the bottom of a panel that is
        ' not there. Nested IFs, never OR(): OR evaluates every argument, and
        ' on a spacer slot the inner INDEX reaches outside the input block.
        FxValue = "=IF(" & sep & "=1,NA(),IF(" & value & "=" & EMPTYTXT & _
                  ",NA()," & placed & "))"
    ElseIf PS_StacksElements(sp) Then
        ' AREA. No separator guard at all - the lead and spacer slots
        ' deliberately hold a neighbouring value so the fill has no wedge to
        ' make. blkx going negative is the only real absence.
        FxValue = "=IF(OR(" & blkx & "<0),NA()," & placed & ")"
    ElseIf PS_Clustered(sp) Then
        ' COLUMN, BAR. Each element owns one sub-slot and leaves the others NA,
        ' which is what produces the clustered look without asking a stacked
        ' group to also cluster - something Excel will not do.
        FxValue = "=IF(OR(" & sep & "=1," & e & "<>" & element & "),NA()," & _
                  placed & ")"
    Else
        FxValue = "=IF(OR(" & sep & "=1),NA()," & placed & ")"
    End If
End Function

' The invisible series that lifts a band's columns off the axis.
'
' A stacked group accumulates EVERYTHING and all bands share the same slots, so
' a band's base cannot be its absolute position - it has to be the delta from
' the running total underneath:
'
'     top(b)  = zero(b) + sum of that band's heights
'     base(b) = zero(b) - top(b-1),   top(-1) = 0
'
' Bases stay NUMERIC on every row, separators included, so the sheet's running
' total and Excel's agree. Being wrong here produces a chart that looks correct
' at one band and floats every panel at three, with no error anywhere - which
' is why the column stage is gated on a THREE-band grid.
Public Function FxBase(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                       ByVal band As Long, ByVal r As Long, _
                       ByVal aggregateOK As Boolean) As String

    Dim col As Long
    col = LY_BaseCol(ly, band)

    Dim zeroHere As String
    zeroHere = band & "+" & Par(ly, PR_ZOFF)

    If band = 0 Then
        FxBase = "=" & zeroHere
        Exit Function
    End If

    Dim c1 As Long, c2 As Long
    c1 = LY_ValCol(ly, sp, band - 1, 0, 0)
    c2 = LY_ValCol(ly, sp, band - 1, sp.Elements - 1, ly.tones - 1)

    Dim prevTop As String
    prevTop = (band - 1) & "+" & Par(ly, PR_ZOFF) & "+" & _
              FxSumIgnoringNA(col, c1, c2, r, aggregateOK)

    FxBase = "=" & zeroHere & "-(" & prevTop & ")"
End Function

' Sum a band's heights while ignoring the #N/A a gap leaves behind.
'
' AGGREGATE(9,6,..) is what the Python engine uses and what manual-steps.md
' documents. It is native from Excel 2010 and needs no _xlfn. prefix when a
' live Excel is doing the writing - that prefix exists for tools writing a file
' Excel has not yet parsed. The caller probes it once per build anyway, because
' the failure that matters is not an error but Excel accepting the name and
' storing something else.
'
' The fallback is EXACT, not approximate: ISNA()->0 is precisely what "ignore
' errors" means for a sum, and a band holds at most `elements` value columns so
' the expansion stays short. A nearly-equivalent fallback would be worse than
' the bug it works around.
Public Function FxSumIgnoringNA(ByVal fromCol As Long, ByVal c1 As Long, _
                                ByVal c2 As Long, ByVal r As Long, _
                                ByVal aggregateOK As Boolean) As String
    If aggregateOK Then
        FxSumIgnoringNA = "AGGREGATE(9,6," & RefSpan(fromCol, c1, c2, r) & ")"
        Exit Function
    End If

    Dim s As String, c As Long, ref1 As String
    For c = c1 To c2
        ref1 = Ref(fromCol, c, r)
        If Len(s) > 0 Then s = s & "+"
        s = s & "IF(ISNA(" & ref1 & "),0," & ref1 & ")"
    Next c
    FxSumIgnoringNA = "(" & s & ")"
End Function

' Area only. Full band height at every lead and spacer slot, drawn as an opaque
' zero-gap-width column, because a stacked area cannot be gapped and its held-
' flat joins have to be painted over rather than left out.
Public Function FxMask(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                       ByVal r As Long) As String
    Dim p As String, s As String
    p = Ref(ly.cMask, ly.cPos, r)
    s = Ref(ly.cMask, ly.cSub, r)
    FxMask = "=IF(OR(" & p & "<=0," & s & ">=" & Par(ly, PR_BLOCKSLOTS) & _
             ")," & PS_Bands(sp) & ",NA())"
End Function


' ----------------------------------------------------------------------------
'  Pin columns - three per direction
'
'  A pin is a marker at the value with a stem back to the panel's zero line.
'  The stem is a custom Y error bar, and an error bar takes ONE colour for the
'  whole series, so a panel of pins that go both ways cannot be one series: the
'  up-pins and the down-pins are separated here, before they reach the chart,
'  rather than coloured afterwards.
' ----------------------------------------------------------------------------

' The sign test that decides whether this tone owns the point.
Private Function PinMine(ByVal valueExpr As String, ByVal tone As Long) As String
    If tone = 0 Then
        PinMine = valueExpr & ">=0"
    Else
        PinMine = valueExpr & "<0"
    End If
End Function

Public Function FxPinValue(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                           ByVal band As Long, ByVal element As Long, _
                           ByVal tone As Long, ByVal r As Long) As String
    Dim col As Long
    col = LY_ValCol(ly, sp, band, element, tone)

    Dim value As String, placed As String, sep As String
    value = FxValueRef(ly, sp, band, element, col, r)
    placed = FxPlaced(ly, sp, band, value)
    sep = Ref(col, ly.cSep, r)

    FxPinValue = "=IF(" & sep & "=1,NA(),IF(" & value & "=" & EMPTYTXT & _
                 ",NA(),IF(" & PinMine(value, tone) & "," & placed & ",NA())))"
End Function

' Stem length in value-axis units. NOTE this writes 0, not NA(), where the
' series does not own the point: an error-bar range must be entirely numeric,
' and one #N/A makes Excel discard the whole range and say nothing.
Public Function FxPinStem(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                          ByVal band As Long, ByVal element As Long, _
                          ByVal tone As Long, ByVal r As Long) As String
    Dim col As Long
    col = LY_ErrCol(ly, band, element, tone)

    Dim value As String, sep As String, stem As String
    value = FxValueRef(ly, sp, band, element, col, r)
    sep = Ref(col, ly.cSep, r)
    stem = "ABS(" & value & ")/" & Par(ly, PR_SPAN) & "*" & Par(ly, PR_BANDFRAC)

    FxPinStem = "=IF(" & sep & "=1,0,IF(" & value & "=" & EMPTYTXT & _
                ",0,IF(" & PinMine(value, tone) & "," & stem & ",0)))"
End Function

' The text a point's data label is linked to.
'
' TEXT() is applied in the CELL, not as a number format on the label, because a
' linked label shows whatever the cell holds - 31.61057692 and all its digits
' otherwise. And the label is linked back to the value the READER typed, not to
' the plotted number: what the chart plots is a banded position between 0 and
' bands, which means nothing to anyone.
Public Function FxPointLabel(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec, _
                             ByRef st As TPanelStyle, _
                             ByVal band As Long, ByVal element As Long, _
                             ByVal tone As Long, ByVal r As Long) As String
    Dim col As Long
    col = LY_LabCol(ly, band, element, tone)

    Dim value As String, sep As String, txt As String
    value = FxValueRef(ly, sp, band, element, col, r)
    sep = Ref(col, ly.cSep, r)
    txt = "TEXT(" & value & "," & Q(st.labelFormat) & ")"

    If PS_TwoTone(sp) Then
        FxPointLabel = "=IF(" & sep & "=1," & EMPTYTXT & ",IF(" & value & _
                       "=" & EMPTYTXT & "," & EMPTYTXT & ",IF(" & _
                       PinMine(value, tone) & "," & txt & "," & EMPTYTXT & ")))"
    Else
        FxPointLabel = "=IF(" & sep & "=1," & EMPTYTXT & ",IF(" & value & _
                       "=" & EMPTYTXT & "," & EMPTYTXT & "," & txt & "))"
    End If
End Function


' ============================================================================
'  The shared scale
'
'  One vmin/vmax for the whole sheet is what makes the panels comparable, and
'  it is computed FROM the input block, so it grows when the data does. That is
'  the difference between this and twenty copied charts.
' ============================================================================

Public Function FxMinRaw(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec) As String
    Dim src As String
    src = InRng(ly)
    If PS_IncludeZero(sp) Then
        FxMinRaw = "=MIN(0,MIN(" & src & "))"
    Else
        ' A dot says "here", not "this far from zero". Dragging the scale down
        ' to a zero the reader is not being asked about squashes every panel
        ' into the top of its band.
        FxMinRaw = "=MIN(" & src & ")"
    End If
End Function

Public Function FxMaxRaw(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec) As String
    If PS_StacksElements(sp) Then
        ' Area stacks its elements, so the shared scale has to be sized on the
        ' stacked TOTAL, not on the largest individual value.
        FxMaxRaw = "=MAX(" & StackSumRng(ly, sp) & ")"
    Else
        FxMaxRaw = "=MAX(" & InRng(ly) & ")"
    End If
End Function

Public Function FxStep(ByRef ly As TPanelLayout) As String
    FxStep = "=10^INT(LOG10(MAX(0.000000001," & Par(ly, PR_MAXRAW) & "-" & _
             Par(ly, PR_MINRAW) & ")))/2"
End Function

Public Function FxVMin(ByRef ly As TPanelLayout, ByRef sp As TPanelSpec) As String
    If sp.FloorHeadroom = 0 Then
        FxVMin = "=FLOOR(" & Par(ly, PR_MINRAW) & "," & Par(ly, PR_STEP) & ")"
    Else
        ' Pad below vmin, mirroring the pad vmax already carries. Set it when
        ' the panels print a label on a NEGATIVE point: such a label sits below
        ' its marker, and on a scale whose span is set by one large positive
        ' outlier the lowest markers sit a hair above vmin and their labels
        ' print through the separator rule beneath. Nothing errors - the digits
        ' are simply struck through.
        FxVMin = "=FLOOR(" & Par(ly, PR_MINRAW) & "-(" & Par(ly, PR_MAXRAW) & _
                 "-" & Par(ly, PR_MINRAW) & ")*" & NumStr(sp.FloorHeadroom) & _
                 "," & Par(ly, PR_STEP) & ")"
    End If
End Function

Public Function FxVMax(ByRef ly As TPanelLayout) As String
    FxVMax = "=CEILING(" & Par(ly, PR_MAXRAW) & "+(" & Par(ly, PR_MAXRAW) & _
             "-" & Par(ly, PR_MINRAW) & ")*" & NumStr(TITLE_HEADROOM) & "," & _
             Par(ly, PR_STEP) & ")"
End Function

Public Function FxSpan(ByRef ly As TPanelLayout) As String
    FxSpan = "=MAX(0.000000001," & Par(ly, PR_VMAX) & "-" & _
             Par(ly, PR_VMIN) & ")"
End Function

' Where value zero sits inside a band, as a fraction of the band's height.
Public Function FxZeroOffset(ByRef ly As TPanelLayout) As String
    FxZeroOffset = "=(0-" & Par(ly, PR_VMIN) & ")/" & Par(ly, PR_SPAN) & _
                   "*" & Par(ly, PR_BANDFRAC)
End Function


' ============================================================================
'  Annotation formulas
' ============================================================================

' A panel's zero line, or nothing when zero is off the shared scale.
'
' A dot grid whose data never crosses zero draws no baseline: a dot says
' "here", and a rule at a zero nobody is being asked about is a claim the chart
' is not making. Type one negative value into the input block and the rule
' appears, without a rebuild - which is the liveness the whole design rests on.
Public Function FxBaselineY(ByRef ly As TPanelLayout, ByVal band As Long) As String
    FxBaselineY = "=IF(OR(" & Par(ly, PR_VMIN) & ">0," & Par(ly, PR_VMAX) & _
                  "<0),NA()," & band & "+" & Par(ly, PR_ZOFF) & ")"
End Function

' A tick label standing in for the hidden value axis. `f` is 0 at the band's
' floor and 1 at its top.
Public Function FxTickText(ByRef ly As TPanelLayout, ByVal f As Double) As String
    FxTickText = "=TEXT(" & Par(ly, PR_VMIN) & "+(" & Par(ly, PR_VMAX) & "-" & _
                 Par(ly, PR_VMIN) & ")*" & NumStr(f) & "," & Q("#,##0") & ")"
End Function

' A tick's position on the value axis, as a plain number.
Public Function TickValue(ByRef sp As TPanelSpec, ByVal band As Long, _
                          ByVal f As Double) As Double
    TickValue = band + f * sp.BandFrac
End Function

' Where a band rule sits: mid-gap between this band and the one below.
Public Function BandRuleY(ByRef sp As TPanelSpec, ByVal band As Long) As Double
    BandRuleY = band - PS_FloorPad(sp)
End Function
