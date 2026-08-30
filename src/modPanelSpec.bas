Attribute VB_Name = "modPanelSpec"
Option Explicit

' ============================================================================
'  modPanelSpec
'  Geometry for a panel chart drawn as ONE native Excel chart object.
'
'  Pure arithmetic - no Excel, no COM, no worksheet, no AddInStorage. This
'  module compiles and runs with every workbook closed, which is the whole
'  reason it is separate: in a project with no test harness, PanelSelfCheck is
'  the only check that can be run for free, and it is the one that catches
'  off-by-ones before they become a chart nobody can read.
'
'  THE MODEL
'  ---------
'  Excel has no panel chart type. The usual workaround is to build one chart
'  and replicate it once per panel, which leaves the user maintaining R x C
'  chart objects whose axes drift apart the moment the data changes. This
'  engine instead puts the whole grid in a single chart by running two
'  coordinate systems through one plot area:
'
'    * The PRIMARY axes carry the data. The category axis holds each panel
'      block's periods laid end to end, separated by spacer categories, with
'      `lead` spacer categories at the left reserved for tick labels. The value
'      axis runs -FloorPad..Bands and is hidden - it is scaffolding, not an
'      axis anyone reads, and it starts below zero so the rule at the foot of
'      the grid clears the bottom band's lowest tick label. A value v belonging
'      to band b lands at
'
'          b + (v - vmin) / (vmax - vmin) * bandFrac
'
'      so every band is a miniature copy of the same scale, which is exactly
'      what makes the panels comparable.
'
'    * The SECONDARY axes carry everything that is not data: panel dividers,
'      band rules, panel titles and the tick labels that stand in for the
'      hidden value axis. With secondary x pinned to min=0.5, max=NSlots+0.5,
'      primary category i sits at secondary x = i exactly. That equality is the
'      whole trick, and it is why nothing may widen the secondary x range to
'      make room - doing so silently slides every annotation off its category.
'
'  ORIENTATION
'  -----------
'  `bar` transposes the model: categories run up the vertical axis, so panel
'  blocks stack vertically and the value bands run left to right. Everything
'  else is identical, which is why the code speaks of BLOCKS (panel groups
'  along the category axis) and BANDS (value offsets) rather than rows and
'  columns:
'
'    line / column / area / pin / dot:  blocks = grid cols, bands = grid rows
'    bar:                               blocks = grid rows, bands = grid cols
'
'  Every consumer that says "rows" when it means "bands" will be correct for
'  five kinds and wrong for bar. The limits in PanelSpecValidate are written
'  against PS_Bands and PS_NSlots for exactly this reason.
'
'  SUB-SLOTS
'  ---------
'  Line, area, pin and dot draw several elements over the same x positions, so
'  one period is one category slot. Column and bar cannot - two clustered
'  columns need two places to stand - so for those kinds each period expands
'  into `elements` adjacent slots and each element series fills only its own,
'  leaving NA in the others. That produces the clustered look without asking a
'  stacked group to also cluster, which Excel will not do.
'
'  WHY A TYPE AND NOT A CLASS
'  --------------------------
'  The Python this ports from uses a dataclass with computed @property values.
'  A VBA Type cannot hold computed properties, so each one is a module-level
'  PS_* function taking the Type ByRef. A class module would model it properly,
'  but a class exports as .cls and this project ships .bas files plus a
'  paste-in .txt for anything non-importable. One unavoidable paste-in (the
'  form) is a cost; two is a choice.
'
'  Ported from panel_spec.py in the panel-charts skill. Keep the two in step.
' ============================================================================


' ----------------------------------------------------------------------------
'  Kinds
' ----------------------------------------------------------------------------

Public Enum ePanelKind
    pkLine = 1
    pkColumn = 2
    pkArea = 3
    pkBar = 4
    pkPin = 5
    pkDot = 6
End Enum

Public Const PK_FIRST As Long = 1
Public Const PK_LAST As Long = 6


' ----------------------------------------------------------------------------
'  The spec
' ----------------------------------------------------------------------------

Public Type TPanelSpec
    ' --- what the user chose
    rows            As Long
    cols            As Long
    Periods         As Long
    Elements        As Long
    kind            As ePanelKind

    ' --- labels. Dynamic arrays are legal inside a Type; they are filled with
    '     defaults by PanelSpecInit when left empty.
    panelNames()    As String
    periodLabels()  As String
    elementNames()  As String

    ' --- tuning
    BandFrac        As Double       ' share of a band's height the plot occupies
    lead            As Long         ' tick-label margin; 0 on input = auto-size
    NTicks          As Long         ' tick labels per band, both ends included

    ' Python's include_zero=None ("whatever the kind wants") is a tri-state and
    ' VBA has no Nothing for a Boolean, so it is carried in two fields. A Type
    ' initialises to False, so the UNSET case is the DEFAULT case - which is
    ' what we want: zeroSet=False means "ask the kind".
    includeZero     As Boolean
    zeroSet         As Boolean

    ' Pad below vmin, as a share of the data's own range, mirroring the pad
    ' vmax already carries. Zero by default, because a grid whose data starts
    ' at zero should still start at zero.
    '
    ' Set it when the panels print a value label on a NEGATIVE point. Such a
    ' label is placed below its marker, so on a scale whose span is set by one
    ' large positive outlier the lowest markers sit a hair above vmin and their
    ' labels overflow the band and print through the separator rule beneath.
    ' Nothing errors; the digits are simply struck through. 0.10 is the value
    ' measured against a grid that renders cleanly.
    FloorHeadroom   As Double

    SheetName       As String
    title           As String
End Type


' ----------------------------------------------------------------------------
'  Layout invariants and limits
'
'  These are deliberately NOT in tblConstants. The test for a settings-sheet
'  constant is "would a user ever want a different value, and does the code
'  still work if they pick one?" - for the limits below the answer is that they
'  are derived from the chart's own size, and PanelSetPlotExtent is how that
'  gets in without this module ever touching a worksheet.
' ----------------------------------------------------------------------------

' Outer bound on the grid, all kinds. A plain, explainable cap that sits on top
' of the kind-aware rules below and never fights them.
Private Const MAX_ROWS As Long = 8
Private Const MAX_COLS As Long = 6

' Bands share the chart's other dimension, one panel deep each.
Private Const MAX_BANDS As Long = 8
Private Const WARN_BANDS As Long = 7

' Points per category slot. Below 2pt a mark is not a mark; below 4pt it is
' legible but crowded. Measured against the reference grids, not guessed.
Private Const MIN_PT_PER_SLOT_HARD As Double = 2#
Private Const MIN_PT_PER_SLOT_WARN As Double = 4#

' Default plot extent in points, less the title, legend and axis-label
' allowance. A slot runs ACROSS the width for the vertical kinds and UP the
' height for bar, which is why the two are held separately.
Private Const DEF_PLOT_W_PT As Double = 720#
Private Const DEF_PLOT_H_PT As Double = 400#

Private mPlotW As Double
Private mPlotH As Double


' ============================================================================
'  Init and validation
' ============================================================================

' Fill in everything the user did not supply, and normalise what they did.
'
' Returns "" when the spec is usable, or the first problem in plain English.
' The form calls this on every keystroke and the Immediate-window entry point
' calls it once - the SAME function, so the rules have no second place to
' drift from.
Public Function PanelSpecInit(ByRef sp As TPanelSpec) As String

    Dim problem As String
    Dim i As Long

    ' --- kind first: every other default depends on it
    If sp.kind < PK_FIRST Or sp.kind > PK_LAST Then
        PanelSpecInit = "Chart kind is not one of the six known kinds."
        Exit Function
    End If

    ' --- the four shape numbers
    problem = AtLeastOne(sp.rows, "Rows")
    If Len(problem) = 0 Then problem = AtLeastOne(sp.cols, "Columns")
    If Len(problem) = 0 Then problem = AtLeastOne(sp.Periods, "Periods")
    If Len(problem) = 0 Then problem = AtLeastOne(sp.Elements, "Elements")
    If Len(problem) > 0 Then
        PanelSpecInit = problem
        Exit Function
    End If

    ' --- tuning defaults, applied before anything reads them
    If sp.BandFrac = 0 Then sp.BandFrac = 0.82
    If sp.NTicks = 0 Then sp.NTicks = 3
    If Len(sp.SheetName) = 0 Then sp.SheetName = "Panel"

    If sp.BandFrac < 0.1 Or sp.BandFrac > 0.98 Then
        PanelSpecInit = "Band fraction must be between 0.1 and 0.98."
        Exit Function
    End If
    If sp.NTicks < 2 Then
        PanelSpecInit = "Ticks per band must be at least 2 - both ends count."
        Exit Function
    End If

    ' --- the tick-label margin
    '
    ' The margin has to be measured in CATEGORIES, but the label it holds has a
    ' fixed width in points - so the wider the grid, the narrower each slot and
    ' the more of them one label spans. A fixed lead is therefore too wide for
    ' a small grid and too narrow for a clustered column grid, whose slots are
    ' `elements` times denser.
    If sp.lead = 0 Then
        Dim span As Long
        span = PS_Blocks(sp) * PS_BlockSlots(sp)
        sp.lead = Round(span * 0.07)
        If sp.lead < 2 Then sp.lead = 2
        If sp.lead > 8 Then sp.lead = 8
    End If
    If sp.lead < 1 Then
        PanelSpecInit = "Lead must be at least 1 category slot."
        Exit Function
    End If

    ' --- default labels
    If Not IsStrArrayFilled(sp.panelNames) Then
        ReDim sp.panelNames(0 To PS_NPanels(sp) - 1)
        For i = 0 To PS_NPanels(sp) - 1
            sp.panelNames(i) = "Panel " & Format$(i + 1, "00")
        Next i
    End If
    If Not IsStrArrayFilled(sp.periodLabels) Then
        ReDim sp.periodLabels(0 To sp.Periods - 1)
        For i = 0 To sp.Periods - 1
            sp.periodLabels(i) = "P" & (i + 1)
        Next i
    End If
    If Not IsStrArrayFilled(sp.elementNames) Then
        ReDim sp.elementNames(0 To sp.Elements - 1)
        For i = 0 To sp.Elements - 1
            sp.elementNames(i) = "Series " & (i + 1)
        Next i
    End If

    ' --- label counts must match the shape, or the fan-out reads past its data
    problem = CountMatches(sp.panelNames, PS_NPanels(sp), "Panel names")
    If Len(problem) = 0 Then problem = CountMatches(sp.periodLabels, sp.Periods, "Period labels")
    If Len(problem) = 0 Then problem = CountMatches(sp.elementNames, sp.Elements, "Element names")
    If Len(problem) > 0 Then
        PanelSpecInit = problem
        Exit Function
    End If

    PanelSpecInit = ""
End Function


' The size limits, kept apart from PanelSpecInit so the form can report a shape
' that is arithmetically sound but too dense to read, without the two concerns
' bleeding into each other.
'
' Returns "" when the grid is buildable, or the reason it is refused.
'
' NOTE the rules are written against PS_Bands and PS_NSlots, never against rows
' and cols. `bar` transposes the model, so "no more than 8 rows" means panel
' DEPTH for five kinds and stacked category SLOTS for bar - an 8-row line grid
' is 8 panels tall and fine, while an 8-row bar grid with 2 elements is 192
' category slots up a 400pt axis at 2pt per bar. Same numbers, different chart.
Public Function PanelSpecValidate(ByRef sp As TPanelSpec) As String

    Dim problem As String
    problem = PanelSpecInit(sp)
    If Len(problem) > 0 Then
        PanelSpecValidate = problem
        Exit Function
    End If

    If sp.rows > MAX_ROWS Then
        PanelSpecValidate = "Maximum " & MAX_ROWS & " rows."
        Exit Function
    End If
    If sp.cols > MAX_COLS Then
        PanelSpecValidate = "Maximum " & MAX_COLS & " columns."
        Exit Function
    End If

    If PS_Bands(sp) > MAX_BANDS Then
        ' Says "rows" or "columns" depending on the kind, because that is what
        ' the user has in front of them - they did not ask for bands.
        PanelSpecValidate = "Maximum " & MAX_BANDS & " " & BandWordFor(sp) & _
                            " for a " & KindName(sp.kind) & " chart."
        Exit Function
    End If

    Dim pps As Double
    pps = PS_PointsPerSlot(sp)
    If pps < MIN_PT_PER_SLOT_HARD Then
        PanelSpecValidate = _
            PS_NSlots(sp) & " category slots " & SlotDirectionFor(sp) & _
            " - about " & Format$(pps, "0.0") & "pt each, too narrow to draw. " & _
            "Reduce " & SlotDriverFor(sp) & "."
        Exit Function
    End If

    PanelSpecValidate = ""
End Function


' Legibility warnings: the grid will build, but the user should see this first.
' The panel-charts skill is explicit that the depth warning is to be raised
' BEFORE building, not after, so it is returned here rather than shown once the
' chart is on screen.
'
' Returns "" when there is nothing to say.
Public Function PanelSpecWarnings(ByRef sp As TPanelSpec) As String

    Dim msg As String

    Dim pps As Double
    pps = PS_PointsPerSlot(sp)
    If pps < MIN_PT_PER_SLOT_WARN Then
        msg = PS_NSlots(sp) & " category slots " & SlotDirectionFor(sp) & _
              ", about " & Format$(pps, "0.0") & "pt each. Fewer " & _
              SlotDriverFor(sp) & " will read better."
    End If

    If PS_Bands(sp) >= WARN_BANDS And sp.Elements > 1 Then
        If Len(msg) > 0 Then msg = msg & vbCrLf
        msg = msg & "At " & PS_Bands(sp) & " " & BandWordFor(sp) & _
              " each panel is roughly 45px tall. Two series only separate in " & _
              "that space if the gap between them is large relative to the " & _
              "shared scale. If the trend is the message this is fine; if the " & _
              "gap between the two series is the message, use fewer " & _
              BandWordFor(sp) & " or plot the variance itself as a pin grid."
    End If

    PanelSpecWarnings = msg
End Function


' Tell the geometry how big the plot area will be, so the density limits track
' the chart size instead of assuming it.
'
' This module never reads a worksheet - that is what makes PanelSelfCheck
' runnable with Excel empty - so the UI layer reads PANEL_CHART_WIDTH and
' PANEL_CHART_HEIGHT from AddInStorage and passes them in here once.
Public Sub PanelSetPlotExtent(ByVal widthPt As Double, ByVal heightPt As Double)
    If widthPt > 0 Then mPlotW = widthPt
    If heightPt > 0 Then mPlotH = heightPt
End Sub


' ============================================================================
'  Shape
' ============================================================================

Public Function PS_NPanels(ByRef sp As TPanelSpec) As Long
    PS_NPanels = sp.rows * sp.cols
End Function

' True when categories run vertically, i.e. the bar kind.
Public Function PS_Transposed(ByRef sp As TPanelSpec) As Boolean
    PS_Transposed = (sp.kind = pkBar)
End Function

' True when each period expands into one slot per element.
Public Function PS_Clustered(ByRef sp As TPanelSpec) As Boolean
    PS_Clustered = (sp.kind = pkColumn Or sp.kind = pkBar)
End Function

' True when the band offset is carried by an invisible base series.
Public Function PS_Stacked(ByRef sp As TPanelSpec) As Boolean
    PS_Stacked = (sp.kind = pkColumn Or sp.kind = pkArea Or sp.kind = pkBar)
End Function

' True when the ELEMENTS also accumulate, so the shared scale must be sized on
' the stacked total rather than on individual values.
Public Function PS_StacksElements(ByRef sp As TPanelSpec) As Boolean
    PS_StacksElements = (sp.kind = pkArea)
End Function

' True when each element needs a series per direction of variance.
'
' An Excel error bar takes one colour for the whole series, and a pin's stem IS
' an error bar, so a panel of pins that go both ways cannot be one series - the
' up-pins and the down-pins have to be separated before they reach the chart,
' not coloured afterwards.
Public Function PS_TwoTone(ByRef sp As TPanelSpec) As Boolean
    PS_TwoTone = (sp.kind = pkPin)
End Function

' True when each point's own value has to be printed beside it.
'
' A column has a length a reader can measure and a line has a path they can
' follow; a pin head and a dot have neither.
Public Function PS_Labelled(ByRef sp As TPanelSpec) As Boolean
    PS_Labelled = (sp.kind = pkPin Or sp.kind = pkDot)
End Function

' Spacer categories between consecutive panel blocks.
'
' Two for area, because a stacked area cannot be gapped and its join has to be
' held flat and then masked; two for bar, because a panel title sits in the
' spacer and one slot is not enough clearance for it.
Public Function PS_Spacers(ByRef sp As TPanelSpec) As Long
    If sp.kind = pkArea Or sp.kind = pkBar Then
        PS_Spacers = 2
    Else
        PS_Spacers = 1
    End If
End Function

' Spare categories past the last block.
'
' Only bar needs them. Its panel titles sit above their block on the CATEGORY
' axis, and the topmost block otherwise ends flush against the axis maximum
' with nowhere for its title to go. The vertical kinds put titles on the value
' axis, which always has band headroom.
Public Function PS_Trailing(ByRef sp As TPanelSpec) As Long
    If PS_Transposed(sp) Then PS_Trailing = PS_Spacers(sp) Else PS_Trailing = 0
End Function

' True when period 1 must be written to the LAST slot of a block.
'
' Excel draws category 1 at the bottom of a bar chart, so a bar panel would
' read P1 at the foot and P12 at the head. Reversing the fan-out fixes the
' reading order without touching the axis - reversing the axis instead would
' invert the secondary-axis alignment that every annotation depends on.
Public Function PS_ReversePeriods(ByRef sp As TPanelSpec) As Boolean
    PS_ReversePeriods = (sp.kind = pkBar)
End Function

' Panel groups laid along the category axis.
Public Function PS_Blocks(ByRef sp As TPanelSpec) As Long
    If PS_Transposed(sp) Then PS_Blocks = sp.rows Else PS_Blocks = sp.cols
End Function

' Value-axis bands, one per panel group across the other dimension.
Public Function PS_Bands(ByRef sp As TPanelSpec) As Long
    If PS_Transposed(sp) Then PS_Bands = sp.cols Else PS_Bands = sp.rows
End Function

Public Function PS_Subslots(ByRef sp As TPanelSpec) As Long
    If PS_Clustered(sp) Then PS_Subslots = sp.Elements Else PS_Subslots = 1
End Function

' Category slots consumed by one panel block.
Public Function PS_BlockSlots(ByRef sp As TPanelSpec) As Long
    PS_BlockSlots = sp.Periods * PS_Subslots(sp)
End Function

' Total categories: the lead margin, every block, the spacers between
' consecutive blocks, and bar's trailing gap.
Public Function PS_NSlots(ByRef sp As TPanelSpec) As Long
    PS_NSlots = sp.lead _
              + PS_Blocks(sp) * PS_BlockSlots(sp) _
              + (PS_Blocks(sp) - 1) * PS_Spacers(sp) _
              + PS_Trailing(sp)
End Function

' Distance from one block's start to the next.
Public Function PS_BlockPitch(ByRef sp As TPanelSpec) As Long
    PS_BlockPitch = PS_BlockSlots(sp) + PS_Spacers(sp)
End Function

' Data series only; the annotation layer adds its own, and how many depends on
' how the chunker splits them - which is why PanelVerify counts what was
' actually emitted rather than trusting a number computed here.
Public Function PS_NSeries(ByRef sp As TPanelSpec) As Long
    Dim perBand As Long
    perBand = sp.Elements
    If PS_TwoTone(sp) Then perBand = perBand * 2
    If PS_Stacked(sp) Then perBand = perBand + 1
    PS_NSeries = PS_Bands(sp) * perBand
End Function

' Whether the shared scale is dragged to include zero.
'
' A dot says "here", not "this far from zero", and dragging the scale down to a
' zero the reader is not being asked about squashes every panel into the top of
' its band. An explicit choice always wins.
Public Function PS_IncludeZero(ByRef sp As TPanelSpec) As Boolean
    If sp.zeroSet Then
        PS_IncludeZero = sp.includeZero
    Else
        PS_IncludeZero = (sp.kind <> pkDot)
    End If
End Function


' ============================================================================
'  Slot maths
' ============================================================================

' 1-based category slot where a block's first period begins.
Public Function PS_BlockStart(ByRef sp As TPanelSpec, ByVal blk As Long) As Long
    PS_BlockStart = sp.lead + blk * PS_BlockPitch(sp) + 1
End Function

Public Function PS_BlockEnd(ByRef sp As TPanelSpec, ByVal blk As Long) As Long
    PS_BlockEnd = PS_BlockStart(sp, blk) + PS_BlockSlots(sp) - 1
End Function

' First spacer category following a block; only blocks 0 .. Blocks-2.
Public Function PS_SpacerFirst(ByRef sp As TPanelSpec, ByVal blk As Long) As Long
    PS_SpacerFirst = PS_BlockEnd(sp, blk) + 1
End Function

Public Function PS_SpacerLast(ByRef sp As TPanelSpec, ByVal blk As Long) As Long
    PS_SpacerLast = PS_SpacerFirst(sp, blk) + PS_Spacers(sp) - 1
End Function

' Middle of the gap after a block - where a divider belongs.
Public Function PS_SpacerCentre(ByRef sp As TPanelSpec, ByVal blk As Long) As Double
    PS_SpacerCentre = (PS_SpacerFirst(sp, blk) + PS_SpacerLast(sp, blk)) / 2#
End Function

' Category slot for one datum. Elements share a slot unless clustered.
'
' The result is ALREADY 1-BASED, because a category slot is a category number
' and Excel numbers chart points from 1. Every caller that feeds a Points()
' index takes it straight from here; adding one walks every mark and every
' label one period to the right, which presents as labels that come and go
' rather than as an error, and PanelVerify passes throughout.
Public Function PS_Slot(ByRef sp As TPanelSpec, ByVal blk As Long, _
                        ByVal period As Long, _
                        Optional ByVal element As Long = 0) As Long
    Dim k As Long
    If PS_ReversePeriods(sp) Then k = sp.Periods - 1 - period Else k = period

    Dim off As Long
    If PS_Clustered(sp) Then off = element Else off = 0

    PS_Slot = PS_BlockStart(sp, blk) + k * PS_Subslots(sp) + off
End Function

' Middle of a period's slots - where its axis label belongs.
Public Function PS_PeriodCentre(ByRef sp As TPanelSpec, ByVal blk As Long, _
                                ByVal period As Long) As Double
    PS_PeriodCentre = PS_Slot(sp, blk, period, 0) + (PS_Subslots(sp) - 1) / 2#
End Function


' ============================================================================
'  Grid maths
' ============================================================================

' Map a grid cell to (block, band).
'
' Bands are numbered from the value axis origin: upward for the vertical kinds,
' rightward for bar. Grid row 0 is the TOP row, so the vertical kinds reverse
' it; bar reverses the block instead, because Excel draws category 1 at the
' bottom of a bar chart.
Public Sub PS_Place(ByRef sp As TPanelSpec, ByVal r As Long, ByVal c As Long, _
                    ByRef blk As Long, ByRef band As Long)
    If PS_Transposed(sp) Then
        blk = sp.rows - 1 - r
        band = c
    Else
        blk = c
        band = sp.rows - 1 - r
    End If
End Sub

' Row-major panel number, matching the input block's row order.
Public Function PS_PanelIndex(ByRef sp As TPanelSpec, ByVal r As Long, _
                              ByVal c As Long) As Long
    PS_PanelIndex = r * sp.cols + c
End Function

' The O(1) inverse of PS_Place: which panel sits at (block, band).
'
' Returns -1 when the pair is outside the grid, so a caller can loop blocks x
' bands without first working out which combinations exist.
Public Function PS_PanelAt(ByRef sp As TPanelSpec, ByVal blk As Long, _
                           ByVal band As Long) As Long
    Dim r As Long, c As Long

    If PS_Transposed(sp) Then
        r = sp.rows - 1 - blk
        c = band
    Else
        r = sp.rows - 1 - band
        c = blk
    End If

    If r < 0 Or r >= sp.rows Or c < 0 Or c >= sp.cols Then
        PS_PanelAt = -1
    Else
        PS_PanelAt = PS_PanelIndex(sp, r, c)
    End If
End Function


' ============================================================================
'  Band maths
' ============================================================================

' Value-axis position of a band's floor.
Public Function PS_BandBase(ByRef sp As TPanelSpec, ByVal band As Long) As Double
    PS_BandBase = CDbl(band)
End Function

' Where the band's plotted content stops, below the next band's floor.
Public Function PS_BandTop(ByRef sp As TPanelSpec, ByVal band As Long) As Double
    PS_BandTop = band + sp.BandFrac
End Function

' How far the value axis runs below the bottom band's floor.
'
' Every band has a gap above it - 1 - bandFrac of a band - and the band's own
' title sits in it. The bottom band had nothing below it: its lowest tick label
' is centred on value 0 and the axis floor was value 0 too, so the baseline
' rule printed straight through the digits. Half a band gap is enough to clear
' an 8pt label at any grid size this builds, and it is the same half-gap the
' separators already use.
Public Function PS_FloorPad(ByRef sp As TPanelSpec) As Double
    PS_FloorPad = (1# - sp.BandFrac) / 2#
End Function

' Points of chart available to one category slot.
'
' A slot is a category. For the vertical kinds they run ACROSS the chart's
' width; for bar they run UP its height, because bar transposes the model. Same
' arithmetic, different extent - and this asymmetry is the whole reason the
' limits are written against PS_NSlots rather than against rows and cols.
Public Function PS_PointsPerSlot(ByRef sp As TPanelSpec) As Double
    Dim extent As Double
    If PS_Transposed(sp) Then extent = PlotH() Else extent = PlotW()
    PS_PointsPerSlot = extent / CDbl(PS_NSlots(sp))
End Function


' ============================================================================
'  Naming
' ============================================================================

Public Function KindName(ByVal k As ePanelKind) As String
    Select Case k
        Case pkLine:   KindName = "line"
        Case pkColumn: KindName = "column"
        Case pkArea:   KindName = "area"
        Case pkBar:    KindName = "bar"
        Case pkPin:    KindName = "pin"
        Case pkDot:    KindName = "dot"
        Case Else:     KindName = "?"
    End Select
End Function

' For the form's dropdown, which shows them capitalised.
Public Function KindLabel(ByVal k As ePanelKind) As String
    Dim s As String
    s = KindName(k)
    KindLabel = UCase$(Left$(s, 1)) & Mid$(s, 2)
End Function

' A one-line summary of what a spec will actually build. The form shows this
' live on every keystroke; it costs nothing because nothing here touches COM.
Public Function PanelSpecDescribe(ByRef sp As TPanelSpec) As String
    PanelSpecDescribe = _
        sp.rows & " x " & sp.cols & " = " & PS_NPanels(sp) & " panels" & _
        "  |  " & sp.Periods & " periods" & _
        "  |  " & sp.Elements & " element" & IIf(sp.Elements = 1, "", "s") & _
        "  |  " & PS_NSlots(sp) & " category slots at " & _
        Format$(PS_PointsPerSlot(sp), "0.0") & "pt" & _
        "  |  " & PS_NSeries(sp) & " data series"
End Function


' ============================================================================
'  Self-check
'
'  Assert the slot arithmetic closes. Cheap, needs no workbook, and it has
'  caught off-by-ones. Returns "" on pass, or the first failure WITH the shape
'  that produced it - a bare "assertion failed" would be no better than
'  nothing.
'
'      Immediate window:   ?modPanelSpec.PanelSelfCheck
'
'  This doubles as the install smoke test: it fails loudly if this module did
'  not import.
' ============================================================================

Public Function PanelSelfCheck() As String

    Dim shapes As Variant
    ' The four shapes the Python carries, plus one it does not: 5x4x6x2 is the
    ' worst case for the render cap - 60 baseline points against 29 categories,
    ' three chunks. The baseline truncation SHIPPED in the Python deliverable
    ' precisely because verify() counted series but never asked how many points
    ' one of them would draw.
    shapes = Array(Array(5, 4, 12, 1), _
                   Array(2, 3, 12, 2), _
                   Array(1, 1, 4, 3), _
                   Array(4, 4, 6, 2), _
                   Array(5, 4, 6, 2))

    Dim k As Long, i As Long, problem As String

    For k = PK_FIRST To PK_LAST
        For i = LBound(shapes) To UBound(shapes)
            problem = CheckOneShape(CLng(k), _
                                    CLng(shapes(i)(0)), CLng(shapes(i)(1)), _
                                    CLng(shapes(i)(2)), CLng(shapes(i)(3)))
            If Len(problem) > 0 Then
                PanelSelfCheck = problem
                Exit Function
            End If
        Next i
    Next k

    ' The transposition case. 8 rows x 4 cols x 12 periods x 2 elements is a
    ' perfectly good LINE grid - 8 bands deep - and an unbuildable BAR grid,
    ' because bar turns those 8 rows into 8 stacked blocks of 24 category slots
    ' each. If these two ever agree, the limits have stopped respecting the
    ' transposition and every other kind-aware rule is worth re-checking.
    problem = CheckTransposeLimits()
    If Len(problem) > 0 Then
        PanelSelfCheck = problem
        Exit Function
    End If

    PanelSelfCheck = "panel geometry self-check passed" & _
                     " (" & (PK_LAST - PK_FIRST + 1) & " kinds x " & _
                     (UBound(shapes) - LBound(shapes) + 1) & " shapes)"
End Function


Private Function CheckOneShape(ByVal kind As Long, ByVal rows As Long, _
                               ByVal cols As Long, ByVal Periods As Long, _
                               ByVal Elements As Long) As String

    Dim sp As TPanelSpec
    Dim tag As String
    Dim problem As String

    sp.kind = kind
    sp.rows = rows: sp.cols = cols
    sp.Periods = Periods: sp.Elements = Elements

    tag = " [" & KindName(kind) & " " & rows & "x" & cols & _
          " p" & Periods & " e" & Elements & "]"

    problem = PanelSpecInit(sp)
    If Len(problem) > 0 Then
        CheckOneShape = "init refused a legal shape: " & problem & tag
        Exit Function
    End If

    Dim b As Long

    ' Every block ends before the next one starts, with the spacers exactly
    ' filling the gap between them - no overlap, no unclaimed slot.
    For b = 0 To PS_Blocks(sp) - 2
        If PS_SpacerLast(sp, b) - PS_SpacerFirst(sp, b) + 1 <> PS_Spacers(sp) Then
            CheckOneShape = "spacer run is the wrong length after block " & b & tag
            Exit Function
        End If
        If PS_SpacerFirst(sp, b) <> PS_BlockEnd(sp, b) + 1 Then
            CheckOneShape = "gap after block " & b & " does not start where the block ends" & tag
            Exit Function
        End If
        If PS_SpacerLast(sp, b) <> PS_BlockStart(sp, b + 1) - 1 Then
            CheckOneShape = "gap after block " & b & " does not end where block " & _
                            (b + 1) & " starts" & tag
            Exit Function
        End If
    Next b

    ' The last block ends where the trailing gap begins.
    If PS_BlockEnd(sp, PS_Blocks(sp) - 1) <> PS_NSlots(sp) - PS_Trailing(sp) Then
        CheckOneShape = "last block ends at " & PS_BlockEnd(sp, PS_Blocks(sp) - 1) & _
                        " but the table runs to " & (PS_NSlots(sp) - PS_Trailing(sp)) & tag
        Exit Function
    End If

    ' Slots are unique and inside their own block. A duplicate here is two
    ' series silently sharing one category.
    Dim seen() As Boolean
    ReDim seen(1 To PS_NSlots(sp))

    Dim p As Long, e As Long, x As Long
    For b = 0 To PS_Blocks(sp) - 1
        For p = 0 To sp.Periods - 1
            For e = 0 To PS_Subslots(sp) - 1
                x = PS_Slot(sp, b, p, e)
                If x < PS_BlockStart(sp, b) Or x > PS_BlockEnd(sp, b) Then
                    CheckOneShape = "slot " & x & " for block " & b & " period " & p & _
                                    " element " & e & " is outside its block" & tag
                    Exit Function
                End If
                If seen(x) Then
                    CheckOneShape = "slot " & x & " is claimed twice" & tag
                    Exit Function
                End If
                seen(x) = True
            Next e
        Next p
    Next b

    ' Every grid cell lands on a distinct (block, band), and PS_PanelAt really
    ' is the inverse of PS_Place.
    Dim used() As Boolean
    ReDim used(0 To PS_Blocks(sp) - 1, 0 To PS_Bands(sp) - 1)

    Dim r As Long, c As Long, blk As Long, band As Long
    For r = 0 To sp.rows - 1
        For c = 0 To sp.cols - 1
            PS_Place sp, r, c, blk, band
            If blk < 0 Or blk >= PS_Blocks(sp) Or band < 0 Or band >= PS_Bands(sp) Then
                CheckOneShape = "cell (" & r & "," & c & ") maps outside the grid" & tag
                Exit Function
            End If
            If used(blk, band) Then
                CheckOneShape = "two cells map to block " & blk & " band " & band & tag
                Exit Function
            End If
            used(blk, band) = True

            If PS_PanelAt(sp, blk, band) <> PS_PanelIndex(sp, r, c) Then
                CheckOneShape = "PS_PanelAt is not the inverse of PS_Place at (" & _
                                r & "," & c & ")" & tag
                Exit Function
            End If
        Next c
    Next r

    ' The render cap. An XY scatter overlaid on a category-axis chart is
    ' TRUNCATED to the number of categories, so no annotation family may be
    ' emitted as a single series longer than NSlots. The baseline is the one
    ' that outgrows it: 3 points per panel, blocks x bands panels.
    problem = CheckChunking(sp, 3 * PS_Blocks(sp) * PS_Bands(sp), 3, tag)
    If Len(problem) > 0 Then
        CheckOneShape = problem
        Exit Function
    End If

    CheckOneShape = ""
End Function


' Prove the chunk arithmetic never exceeds the cap and never cuts inside a
' segment. A cut through the middle of a three-point polyline loses that
' segment's line entirely - and nothing errors.
Private Function CheckChunking(ByRef sp As TPanelSpec, ByVal nPoints As Long, _
                               ByVal unit As Long, ByVal tag As String) As String

    Dim cap As Long
    cap = PS_NSlots(sp)

    Dim per As Long
    per = PanelChunkSize(cap, unit)

    If per < unit Then
        CheckChunking = "chunk size " & per & " is smaller than one segment" & tag
        Exit Function
    End If
    If per Mod unit <> 0 Then
        CheckChunking = "chunk size " & per & " would cut inside a segment" & tag
        Exit Function
    End If

    Dim taken As Long, thisChunk As Long, guard As Long
    Do While taken < nPoints
        If taken + per > nPoints Then thisChunk = nPoints - taken Else thisChunk = per
        If thisChunk > cap Then
            CheckChunking = "a chunk of " & thisChunk & " exceeds the " & cap & _
                            " point render cap" & tag
            Exit Function
        End If
        taken = taken + thisChunk

        guard = guard + 1
        If guard > 10000 Then
            CheckChunking = "chunking did not terminate" & tag
            Exit Function
        End If
    Loop

    CheckChunking = ""
End Function


' Points per chunk: as many whole segments as fit under the cap, and never
' fewer than one segment even if a single segment is itself over the cap -
' there is nothing useful to do in that case but draw it.
'
' Integer division. "cap / unit" returns a Double, and a Double flowing into a
' chunk count is a rounding bug waiting for a large grid.
Public Function PanelChunkSize(ByVal cap As Long, ByVal unit As Long) As Long
    Dim u As Long
    u = unit
    If u < 1 Then u = 1

    PanelChunkSize = (cap \ u) * u
    If PanelChunkSize < u Then PanelChunkSize = u
End Function


Private Function CheckTransposeLimits() As String

    Dim spLine As TPanelSpec, spBar As TPanelSpec
    Dim msgLine As String, msgBar As String

    spLine.kind = pkLine
    spLine.rows = 8: spLine.cols = 4: spLine.Periods = 12: spLine.Elements = 2
    msgLine = PanelSpecValidate(spLine)

    spBar.kind = pkBar
    spBar.rows = 8: spBar.cols = 4: spBar.Periods = 12: spBar.Elements = 2
    msgBar = PanelSpecValidate(spBar)

    If Len(msgLine) > 0 Then
        CheckTransposeLimits = "8x4 p12 e2 should be a legal LINE grid but was " & _
                               "refused: " & msgLine
        Exit Function
    End If
    If Len(msgBar) = 0 Then
        CheckTransposeLimits = "8x4 p12 e2 was accepted as a BAR grid. Bar " & _
                               "transposes the model, so those 8 rows are " & _
                               PS_NSlots(spBar) & " stacked category slots at " & _
                               Format$(PS_PointsPerSlot(spBar), "0.0") & _
                               "pt each - the limits have stopped respecting it."
        Exit Function
    End If

    CheckTransposeLimits = ""
End Function


' ============================================================================
'  Private helpers
' ============================================================================

Private Function PlotW() As Double
    If mPlotW <= 0 Then mPlotW = DEF_PLOT_W_PT
    PlotW = mPlotW
End Function

Private Function PlotH() As Double
    If mPlotH <= 0 Then mPlotH = DEF_PLOT_H_PT
    PlotH = mPlotH
End Function

Private Function AtLeastOne(ByVal v As Long, ByVal label As String) As String
    If v < 1 Then AtLeastOne = label & " must be at least 1."
End Function

' True when a dynamic String array has actually been sized. LBound on an
' unallocated array raises, which is the only way to tell in VBA.
Private Function IsStrArrayFilled(ByRef a() As String) As Boolean
    Dim n As Long
    On Error Resume Next
    n = UBound(a)
    IsStrArrayFilled = (Err.Number = 0)
    Err.Clear
End Function

Private Function CountMatches(ByRef a() As String, ByVal want As Long, _
                              ByVal label As String) As String
    Dim n As Long
    n = UBound(a) - LBound(a) + 1
    If n <> want Then
        CountMatches = label & ": " & n & " supplied, " & want & " expected."
    End If
End Function

' "rows" or "columns", whichever the bands actually are for this kind. The user
' did not ask for bands, so the message should not mention them.
Private Function BandWordFor(ByRef sp As TPanelSpec) As String
    If PS_Transposed(sp) Then BandWordFor = "columns" Else BandWordFor = "rows"
End Function

Private Function SlotDirectionFor(ByRef sp As TPanelSpec) As String
    If PS_Transposed(sp) Then
        SlotDirectionFor = "stacked up the chart"
    Else
        SlotDirectionFor = "across the chart"
    End If
End Function

' Which number the user should reduce to get fewer slots. For the vertical
' kinds the slots are driven by columns; for bar, by rows.
Private Function SlotDriverFor(ByRef sp As TPanelSpec) As String
    If PS_Transposed(sp) Then
        SlotDriverFor = "rows, periods or elements"
    Else
        SlotDriverFor = "columns, periods or elements"
    End If
End Function
