Attribute VB_Name = "modPanelChart"
Option Explicit

' ============================================================================
'  modPanelChart
'  Turns the engine modPanelBuilder wrote into one chart object, and then
'  measures what Excel actually did.
'
'  THE WHOLE GRID IS ONE CHART, RUNNING TWO COORDINATE SYSTEMS
'  ----------------------------------------------------------
'  The PRIMARY axes carry the data: the category axis holds every panel block's
'  periods laid end to end and the value axis is hidden scaffolding running
'  -floorPad..bands. The SECONDARY axes carry everything else - dividers, band
'  rules, per-panel baselines, panel titles, tick labels - as XY scatter series
'  positioned in category space.
'
'  ORDERING MATTERS AND VBA WILL NOT TELL YOU
'  ------------------------------------------
'  Axes(xlValue, xlSecondary) DOES NOT EXIST until some series has
'  AxisGroup = 2. So the annotation series go on BEFORE SetAxes, exactly as the
'  Python orders it. Reorder them and Excel raises 1004 with a message that
'  says nothing about ordering.
'
'  Ported from panel_excel.py (_draw_chart, _add_*_series, _axes, _legend,
'  _page_setup, verify) in the panel-charts skill. Keep the two in step.
' ============================================================================

' Rules and furniture. Grey rather than black throughout: this layer is
' scaffolding the reader should be able to see past, not content.
Private Const GREY_RULE As Long = 12566463      ' RGB(191,191,191) separators
Private Const GREY_BASE As Long = 8421504       ' RGB(128,128,128) baselines
Private Const GREY_AXIS As Long = 10921638      ' RGB(166,166,166) category axis
Private Const MASK_WHITE As Long = 16777215     ' RGB(255,255,255) the area mask

Private Const DEF_CHART_W As Double = 720
Private Const DEF_CHART_H As Double = 500

' How many series the annotation layer actually cost. Counted from what was
' emitted, never from a literal: the Python's verify() once expected five
' families after a sixth was added, and reported a good grid as broken.
Private mAnnotationSeries As Long


' ============================================================================
'  Draw
' ============================================================================

' Build the chart. `vals` is the same input array the builder was given; it is
' read only for the per-point work a pin or dot grid needs.
Public Function PanelDrawChart(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                               ByRef st As TPanelStyle, ByRef ly As TPanelLayout, _
                               ByRef anns() As TAnnotation, ByVal anCount As Long, _
                               ByRef vals As Variant, _
                               Optional ByVal chartW As Double = 0, _
                               Optional ByVal chartH As Double = 0) As ChartObject

    Dim obj As ChartObject, ch As Chart, sc As SeriesCollection
    Dim anchor As Range, catRng As Range
    Dim band As Long, e As Long, tone As Long, g As Long
    Dim ser As Series

    If chartW <= 0 Then chartW = DEF_CHART_W
    If chartH <= 0 Then chartH = DEF_CHART_H

    mAnnotationSeries = 0

    Set anchor = ws.Cells(ly.chartRow, 1)
    Set obj = ws.ChartObjects.Add(anchor.Left, anchor.Top, chartW, chartH)
    obj.Placement = xlFreeFloating
    Set ch = obj.Chart

    ' Excel will sometimes seed a new chart from whatever the user had
    ' selected, and this tool is very often run WITH a data range selected.
    ' Nothing downstream expects a series it did not add: an extra one breaks
    ' the legend-entry arithmetic and the verify count. The Python never sees
    ' this because it builds into a workbook it created.
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    ch.ChartType = BaseChartType(sp)
    ' xlNotPlotted is 1. xlZero is 2, and plotting a gap as zero drops a spike
    ' to the baseline in every panel. Named constants make this one free.
    ch.DisplayBlanksAs = xlNotPlotted

    Set catRng = ws.Range(ws.Cells(ly.ptR0, ly.cCat), ws.Cells(ly.ptR1, ly.cCat))
    Set sc = ch.SeriesCollection

    ' Data series, BOTTOM BAND FIRST, so the stack order matches the series
    ' order and the legend's first entries are the ones worth keeping.
    For band = 0 To PS_Bands(sp) - 1

        If PS_Stacked(sp) Then
            Set ser = sc.NewSeries
            ' Series.Name = "" raises; a single space does not.
            ser.name = " "
            ser.ChartType = BaseChartType(sp)
            ser.XValues = catRng
            ser.Values = PlotCol(ws, ly, LY_BaseCol(ly, band))
            ser.Format.Fill.Visible = msoFalse
            ser.Format.Line.Visible = msoFalse
        End If

        For e = 0 To sp.Elements - 1
            If PS_TwoTone(sp) Then
                For tone = 0 To ly.tones - 1
                    AddPinSeries ws, sp, st, ly, ch, sc, catRng, band, e, tone, vals
                Next tone
            ElseIf sp.kind = pkDot Then
                AddDotSeries ws, sp, st, ly, ch, sc, catRng, band, e, vals
            Else
                AddPlainSeries ws, sp, st, ly, sc, catRng, band, e
            End If
        Next e
    Next band

    ' The area mask. A stacked area cannot be gapped at all - Excel joins
    ' whatever points it is given - so the joins are held flat by the kx/blkx
    ' formulas and then painted over by an opaque column exactly one category
    ' wide, which is what GapWidth = 0 buys.
    If PS_StacksElements(sp) Then
        Set ser = sc.NewSeries
        ser.name = " "
        ser.ChartType = xlColumnClustered
        ser.XValues = catRng
        ser.Values = PlotCol(ws, ly, ly.cMask)
        ser.Format.Fill.Visible = msoTrue
        ser.Format.Fill.ForeColor.RGB = MASK_WHITE
        ser.Format.Line.Visible = msoFalse
    End If

    For g = 1 To ch.ChartGroups.Count
        On Error Resume Next
        If PS_StacksElements(sp) Then
            ch.ChartGroups(g).GapWidth = 0
        Else
            ch.ChartGroups(g).GapWidth = 40
        End If
        Err.Clear
        On Error GoTo 0
    Next g

    ' Annotations BEFORE axes - see the module header.
    AddAnnotationSeries ws, sp, st, ly, ch, sc, anns, anCount
    SetAxes ch, sp
    SetLegend ch, sp

    If st.chartTitle Then
        ch.HasTitle = True
        If Len(sp.title) > 0 Then
            ch.chartTitle.text = sp.title
        Else
            ch.chartTitle.text = KindLabel(sp.kind) & " panel chart - " & _
                                 sp.rows & " x " & sp.cols
        End If
        ch.chartTitle.Font.Size = 12
        ch.chartTitle.Font.bold = True
    Else
        ch.HasTitle = False
    End If

    Set PanelDrawChart = obj
End Function


Private Function BaseChartType(ByRef sp As TPanelSpec) As XlChartType
    Select Case sp.kind
        Case pkLine:   BaseChartType = xlLine
        Case pkColumn: BaseChartType = xlColumnStacked
        Case pkArea:   BaseChartType = xlAreaStacked
        Case pkBar:    BaseChartType = xlBarStacked
        Case pkPin:    BaseChartType = xlLineMarkers
        Case pkDot:    BaseChartType = xlLineMarkers
        Case Else:     BaseChartType = xlLine
    End Select
End Function


Private Function PlotCol(ByVal ws As Worksheet, ByRef ly As TPanelLayout, _
                         ByVal col As Long) As Range
    Set PlotCol = ws.Range(ws.Cells(ly.ptR0, col), ws.Cells(ly.ptR1, col))
End Function


' line, column and area: one series per element per band.
Private Sub AddPlainSeries(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                           ByRef st As TPanelStyle, ByRef ly As TPanelLayout, _
                           ByVal sc As SeriesCollection, ByVal catRng As Range, _
                           ByVal band As Long, ByVal element As Long)

    Dim ser As Series, colour As Long
    Set ser = sc.NewSeries
    ser.name = sp.elementNames(LBound(sp.elementNames) + element)
    ser.ChartType = BaseChartType(sp)
    ser.XValues = catRng
    ser.Values = PlotCol(ws, ly, LY_ValCol(ly, sp, band, element, 0))

    colour = PanelElementColour(st, element)

    If sp.kind = pkLine Then
        ser.Format.Line.ForeColor.RGB = colour
        ' The first element reads as the subject and the rest as reference,
        ' which at 45px per panel is the only distinction that survives.
        If element = 0 Then
            ser.Format.Line.weight = 1.75
        Else
            ser.Format.Line.weight = 1#
        End If
        ser.markerStyle = xlMarkerStyleNone
        ser.Smooth = False
    Else
        ser.Format.Fill.Visible = msoTrue
        ser.Format.Fill.ForeColor.RGB = colour
        ser.Format.Line.Visible = msoFalse
    End If
End Sub


' ----------------------------------------------------------------------------
'  Pin: a marker at the value with a stem back to the panel's zero line.
'
'  Four things here are not in the documentation and each fails silently or
'  unhelpfully. All four are established in calibration/calib_pin.py, and the
'  first is a VBA hazard the Python does not have:
'
'  * Series.ErrorBar is a SUB, so it is a STATEMENT. Writing
'    ser.ErrorBar(a, b, c, d, e) is a VBA syntax error - a parenthesised
'    multi-argument call to a Sub.
'  * EndStyle takes xlNoCap. xlNone POISONS the chart, and every later
'    PlotArea call fails with it.
'  * The error bar's weight is in POINTS, so the stem width is exact rather
'    than a fraction of a category slot - which is the whole reason a pin is an
'    error bar and not a very narrow column. GapWidth caps at 500, bottoming
'    out around a 9px stem where the reference draws 5.
'  * On a line series Points(i).Format.Line is the CONNECTING SEGMENT, not the
'    marker border. The border is MarkerForegroundColor.
' ----------------------------------------------------------------------------
Private Sub AddPinSeries(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                         ByRef st As TPanelStyle, ByRef ly As TPanelLayout, _
                         ByVal ch As Chart, ByVal sc As SeriesCollection, _
                         ByVal catRng As Range, ByVal band As Long, _
                         ByVal element As Long, ByVal tone As Long, _
                         ByRef vals As Variant)

    Dim ser As Series, base As Long, stemColour As Long
    Dim ecol As Long, addr As String, include As Long

    Set ser = sc.NewSeries
    ' Only the first tone carries the element's name; the second would double
    ' every legend entry.
    If tone = 0 Then
        ser.name = sp.elementNames(LBound(sp.elementNames) + element)
    Else
        ser.name = " "
    End If
    ser.ChartType = BaseChartType(sp)
    ser.XValues = catRng
    ser.Values = PlotCol(ws, ly, LY_ValCol(ly, sp, band, element, tone))
    ser.Format.Line.Visible = msoFalse          ' markers only, never a path
    ser.Smooth = False
    ser.markerStyle = MarkerHead(st, xlMarkerStyleSquare)
    ser.markerSize = ClampMarker(st.markerSize)

    base = PanelElementColour(st, element)
    ser.MarkerBackgroundColor = base
    ser.MarkerForegroundColor = base

    ecol = LY_ErrCol(ly, band, element, tone)
    ' A full external A1 reference. R1C1 is not accepted here, and the sheet
    ' name has its own apostrophes doubled so a sheet renamed to "Bob's panel"
    ' still resolves.
    addr = "='" & SheetRef(ws) & "'!" & PlotCol(ws, ly, ecol).Address(True, True)

    If tone = 0 Then
        include = xlErrorBarIncludeMinusValues
    Else
        include = xlErrorBarIncludePlusValues
    End If
    ser.ErrorBar xlY, include, xlErrorBarTypeCustom, addr, addr

    With ser.ErrorBars
        .EndStyle = xlNoCap
        If st.signColoursSet Then
            If tone = 0 Then stemColour = st.signColourUp Else stemColour = st.signColourDown
        Else
            stemColour = base
        End If
        .Format.Line.ForeColor.RGB = stemColour
        .Format.Line.weight = st.stemWeight
    End With

    If st.valueLabels Then
        LabelPoints ws, sp, st, ly, ser, band, element, tone, vals
    End If
End Sub


' ----------------------------------------------------------------------------
'  Dot: a marker, and nothing else.
'
'  A dot is a pin with the stem taken away, and that subtraction is a claim
'  rather than a saving. A stem says "this far from zero"; a dot says only
'  "here", which is what a reader wants when the question is where a value sits
'  against a reference rather than how big it is. It is also why a dot grid
'  defaults to a scale that does NOT include zero - there is no stem for zero
'  to anchor, and dragging the scale down to a zero nobody asked about squashes
'  every panel into the top of its band.
'
'  Losing the stem costs the two-series split as well: an error bar takes one
'  colour for a whole series, which is what forces a pin's directions apart
'  before they reach the chart. With no error bar a single series carries both
'  signs.
' ----------------------------------------------------------------------------
Private Sub AddDotSeries(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                         ByRef st As TPanelStyle, ByRef ly As TPanelLayout, _
                         ByVal ch As Chart, ByVal sc As SeriesCollection, _
                         ByVal catRng As Range, ByVal band As Long, _
                         ByVal element As Long, ByRef vals As Variant)

    Dim ser As Series, base As Long
    Set ser = sc.NewSeries
    ser.name = sp.elementNames(LBound(sp.elementNames) + element)
    ser.ChartType = BaseChartType(sp)
    ser.XValues = catRng
    ser.Values = PlotCol(ws, ly, LY_ValCol(ly, sp, band, element, 0))
    ser.Format.Line.Visible = msoFalse
    ser.Smooth = False
    ser.markerStyle = MarkerHead(st, xlMarkerStyleCircle)
    ser.markerSize = ClampMarker(st.markerSize)

    base = PanelElementColour(st, element)
    ser.MarkerBackgroundColor = base
    ser.MarkerForegroundColor = base

    If st.valueLabels Then
        LabelPoints ws, sp, st, ly, ser, band, element, 0, vals
    End If
End Sub


' The kind's own head unless the caller named one. Set it when two elements
' need telling apart by SHAPE as well as colour, which is the one distinction a
' greyscale print still carries.
Private Function MarkerHead(ByRef st As TPanelStyle, _
                            ByVal dflt As Long) As Long
    If st.markerStyle = 0 Then MarkerHead = dflt Else MarkerHead = st.markerStyle
End Function

' Points. Excel refuses a marker outside 2..72 and raises doing it.
Private Function ClampMarker(ByVal n As Long) As Long
    ClampMarker = n
    If ClampMarker < 2 Then ClampMarker = 2
    If ClampMarker > 72 Then ClampMarker = 72
End Function


' Print each point's own value beside it - but only where this series actually
' owns the point. Anywhere else the marker is absent, and touching it is both
' wasted and, for the label, an error.
Private Sub LabelPoints(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                        ByRef st As TPanelStyle, ByRef ly As TPanelLayout, _
                        ByVal ser As Series, ByVal band As Long, _
                        ByVal element As Long, ByVal tone As Long, _
                        ByRef vals As Variant)

    Dim lcol As Long, blk As Long, panel As Long, k As Long, idx As Long
    Dim v As Variant, pt As Point, dl As DataLabel

    lcol = LY_LabCol(ly, band, element, tone)

    For blk = 0 To PS_Blocks(sp) - 1
        panel = PS_PanelAt(sp, blk, band)
        If panel >= 0 Then
            For k = 0 To sp.Periods - 1
                v = Datum(vals, sp, panel, element, k)
                If Not IsEmpty(v) Then
                    If (CDbl(v) >= 0) = (tone = 0) Or Not PS_TwoTone(sp) Then

                        ' Point n IS category n, and PS_Slot already counts
                        ' from 1. Adding one to it walks every mark and every
                        ' label one period to the right - which shows up as
                        ' labels that come and go, not as an error, and
                        ' PanelVerify passes throughout.
                        idx = PS_Slot(sp, blk, k, 0)

                        Set pt = Nothing
                        On Error Resume Next
                        Set pt = ser.Points(idx)
                        Err.Clear
                        On Error GoTo 0

                        If Not pt Is Nothing Then
                            pt.HasDataLabel = True
                            Set dl = pt.DataLabel
                            dl.formula = "='" & SheetRef(ws) & "'!" & _
                                ws.Cells(ly.ptR0 + idx - 1, lcol).Address(True, True)
                            If CDbl(v) >= 0 Then
                                dl.Position = xlLabelPositionAbove
                            Else
                                dl.Position = xlLabelPositionBelow
                            End If
                        End If
                    End If
                End If
            Next k
        End If
    Next blk

    ' The font ONCE for the series, not four property sets per point. A 5x4 pin
    ' grid is 240 points; this halves the COM traffic and applies to exactly
    ' the labels that exist.
    On Error Resume Next
    ser.DataLabels.Font.Size = st.labelSize
    ser.DataLabels.Font.name = "Arial"
    Err.Clear
    On Error GoTo 0
End Sub


' ============================================================================
'  The annotation layer
' ============================================================================

' Points per chunk: as many WHOLE SEGMENTS as fit under the render cap.
'
' An XY scatter series overlaid on a category-axis chart is TRUNCATED to the
' number of categories. Excel reports every point, caches every point, and
' silently draws only the first n - which on the 5x4 line grid cost the top
' row's last two panels their axis line, with nothing in the file looking
' wrong. A cut through the middle of a three-point polyline would lose that
' segment's line entirely, so cuts land on segment boundaries only.
'
' Integer division: "cap / unit" returns a Double, and a Double flowing into a
' chunk count is a rounding bug waiting for a large grid.
Private Function ChunkSize(ByVal cap As Long, ByVal unit As Long) As Long
    Dim u As Long
    u = unit
    If u < 1 Then u = 1
    ChunkSize = (cap \ u) * u
    If ChunkSize < u Then ChunkSize = u
End Function


Private Sub AddAnnotationSeries(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                                ByRef st As TPanelStyle, ByRef ly As TPanelLayout, _
                                ByVal ch As Chart, ByVal sc As SeriesCollection, _
                                ByRef anns() As TAnnotation, ByVal anCount As Long)

    Dim i As Long, labelPos As Long

    For i = 1 To anCount
        Select Case anns(i).name
            Case "split"
                EmitAnnotation ws, sp, ly, sc, anns(i), st.dividerWeight, _
                               st.dividerColour, 0, 0, False
            Case "divider", "bandrule"
                EmitAnnotation ws, sp, ly, sc, anns(i), 0.75, GREY_RULE, _
                               0, 0, False
            Case "baseline"
                EmitAnnotation ws, sp, ly, sc, anns(i), 0.75, GREY_BASE, _
                               0, 0, False
            Case "ptitle"
                ' Right, not Above. Right centres the text vertically on its
                ' anchor; Above lifts it a full text-height and lands it on the
                ' divider.
                EmitAnnotation ws, sp, ly, sc, anns(i), 0, 0, _
                               xlLabelPositionRight, st.titleSize, True
            Case "ytick"
                If st.valueTicks Then
                    If PS_Transposed(sp) Then
                        labelPos = xlLabelPositionBelow
                    Else
                        labelPos = xlLabelPositionLeft
                    End If
                    EmitAnnotation ws, sp, ly, sc, anns(i), 0, 0, _
                                   labelPos, 8, False
                Else
                    EmitAnnotation ws, sp, ly, sc, anns(i), 0, 0, 0, 0, False
                End If
        End Select
    Next i
End Sub


' One annotation family, sliced into as many series as the render cap needs,
' each labelled as it is made.
Private Sub EmitAnnotation(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                           ByRef ly As TPanelLayout, ByVal sc As SeriesCollection, _
                           ByRef an As TAnnotation, ByVal weight As Double, _
                           ByVal colour As Long, ByVal labelPos As Long, _
                           ByVal labelSize As Long, ByVal bold As Boolean)

    Dim cap As Long, per As Long, taken As Long, take As Long
    Dim r0 As Long, part As Long, ser As Series
    Dim i As Long, offset As Long, pt As Point, dl As DataLabel

    cap = PS_NSlots(sp)
    per = ChunkSize(cap, an.unit)
    taken = 0
    part = 0

    Do While taken < an.n
        take = an.n - taken
        If take > per Then take = per
        r0 = an.r0 + taken
        part = part + 1

        Set ser = sc.NewSeries
        If part = 1 Then ser.name = an.name Else ser.name = an.name & " " & part
        ser.ChartType = xlXYScatter
        ser.XValues = ws.Range(ws.Cells(r0, an.col), ws.Cells(r0 + take - 1, an.col))
        ser.Values = ws.Range(ws.Cells(r0, an.col + 1), _
                              ws.Cells(r0 + take - 1, an.col + 1))
        ' This is what creates the secondary axes, and why SetAxes runs after.
        ser.AxisGroup = 2
        ser.markerStyle = xlMarkerStyleNone

        If weight <= 0 Then
            ser.Format.Line.Visible = msoFalse
        Else
            ser.Format.Line.Visible = msoTrue
            ser.Format.Line.ForeColor.RGB = colour
            ser.Format.Line.weight = weight
        End If

        If an.labCol > 0 And labelPos <> 0 Then
            ser.HasDataLabels = True
            offset = r0 - an.r0
            For i = 0 To take - 1
                Set pt = Nothing
                On Error Resume Next
                Set pt = ser.Points(i + 1)
                Err.Clear
                On Error GoTo 0
                If Not pt Is Nothing Then
                    ' A grid can hold fewer panels than cells. A label linked
                    ' to a blank name cell prints "0", so a cell with no panel
                    ' gets no label rather than a zero where a name should be.
                    If an.name = "ptitle" And Not PanelNamed(sp, offset + i) Then
                        pt.HasDataLabel = False
                    Else
                        pt.HasDataLabel = True
                        Set dl = pt.DataLabel
                        dl.formula = "='" & SheetRef(ws) & "'!" & _
                            ws.Cells(an.r0 + offset + i, an.labCol).Address(True, True)
                        dl.Position = labelPos
                        dl.Font.Size = labelSize
                        dl.Font.bold = bold
                    End If
                End If
            Next i
        End If

        mAnnotationSeries = mAnnotationSeries + 1
        taken = taken + take
    Loop
End Sub


' ============================================================================
'  Axes
' ============================================================================

Private Sub SetAxes(ByVal ch As Chart, ByRef sp As TPanelSpec)

    Dim va As Axis, ca As Axis, secCat As Axis, secVal As Axis
    Dim floorVal As Double

    ' Room below the bottom band. Its lowest tick label is centred on value 0
    ' with nothing underneath it - every other band's gap is above - so the
    ' rule at the foot of the grid had nowhere to sit but on the digits.
    '
    ' ONE local, assigned to BOTH value axes. Never two expressions: the data
    ' is on one axis and every annotation on the other, and a floor on one
    ' alone slides every label off its own data without erroring.
    floorVal = -PS_FloorPad(sp)

    Set va = ch.Axes(xlValue, xlPrimary)            ' hidden scaffolding
    va.MinimumScale = floorVal
    va.MaximumScale = PS_Bands(sp)
    ' Where the category axis crosses, i.e. how high the bottom rule sits.
    ' Excel's default is value 0, which is the bottom band's lowest tick label,
    ' and the rule printed straight through the digits. At the minimum it lands
    ' half a band gap lower - exactly where band 0's separator would be - so
    ' the foot of the grid is ruled like every other band.
    va.Crosses = xlAxisCrossesMinimum
    va.MajorTickMark = xlTickMarkNone
    va.TickLabelPosition = xlTickLabelPositionNone
    HideLine va
    If va.HasMajorGridlines Then va.MajorGridlines.Delete

    Set ca = ch.Axes(xlCategory, xlPrimary)
    ca.TickLabelPosition = xlLow
    ca.MajorTickMark = xlTickMarkNone
    ca.TickLabels.Font.Size = 8
    On Error Resume Next
    ca.Format.Line.ForeColor.RGB = GREY_AXIS
    Err.Clear
    On Error GoTo 0

    ' bar transposes the model, so which Excel axis is the category axis and
    ' which is the value axis swaps with it.
    If PS_Transposed(sp) Then
        Set secCat = ch.Axes(xlValue, xlSecondary)
        Set secVal = ch.Axes(xlCategory, xlSecondary)
    Else
        Set secCat = ch.Axes(xlCategory, xlSecondary)
        Set secVal = ch.Axes(xlValue, xlSecondary)
    End If

    ' THE ALIGNMENT CONTRACT. With secondary x pinned to 0.5 .. NSlots + 0.5,
    ' primary category i sits at secondary x = i exactly. That equality is what
    ' the entire annotation layer stands on, and widening it to make room for
    ' labels silently slides every divider, title and tick off its category.
    ' Reserve `lead` category slots instead - that is what they are for.
    secCat.MinimumScale = 0.5
    secCat.MaximumScale = PS_NSlots(sp) + 0.5

    secVal.MinimumScale = floorVal
    secVal.MaximumScale = PS_Bands(sp)

    secCat.TickLabelPosition = xlTickLabelPositionNone
    secCat.MajorTickMark = xlTickMarkNone
    HideLine secCat
    secVal.TickLabelPosition = xlTickLabelPositionNone
    secVal.MajorTickMark = xlTickMarkNone
    HideLine secVal
End Sub


' Hide an object's outline, whichever of the two ways it supports.
'
' As Object because Axis and Series share no interface. Neither Axis.Format nor
' Axis.Border is reachable on the secondary x axis in every Excel build - both
' raise E_FAIL - and a cosmetic line is not worth failing a build over.
Private Function HideLine(ByVal o As Object) As Boolean
    On Error Resume Next
    o.Format.Line.Visible = msoFalse
    If Err.Number = 0 Then
        HideLine = True
        Exit Function
    End If
    Err.Clear
    o.Border.LineStyle = xlNone
    HideLine = (Err.Number = 0)
    Err.Clear
End Function


' Show one entry per element and nothing else.
'
' Series are added bottom band first, so the first band's element entries are
' the ones to keep: 1..E normally, 2..E+1 when a stacked kind puts its
' invisible base first.
Private Sub SetLegend(ByVal ch As Chart, ByRef sp As TPanelSpec)

    Dim first As Long, last As Long, i As Long

    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionTop
    ch.Legend.Font.Size = 9

    If PS_Stacked(sp) Then first = 2 Else first = 1
    last = first + sp.Elements - 1

    ' IN REVERSE. Deleting forwards shifts the indices under you and takes out
    ' the wrong entries.
    For i = ch.Legend.LegendEntries.Count To 1 Step -1
        If i < first Or i > last Then
            On Error Resume Next
            ch.Legend.LegendEntries(i).Delete
            Err.Clear
            On Error GoTo 0
        End If
    Next i

    ' One element needs no legend to say which one it is.
    If sp.Elements = 1 Then ch.HasLegend = False
End Sub


' ============================================================================
'  Page setup
' ============================================================================

' One landscape page holding exactly the chart, and no gridlines.
'
' The chart sits below every cell the engine uses, so the print area is the
' rectangle under it and contains nothing else - otherwise the cell contents
' would print straight through the chart.
'
' NOTE there is deliberately no Application.PrintCommunication = False around
' this. build\IMPORT.md records that exact optimisation being tried in this
' add-in and SILENTLY CORRUPTING PageSetup writes - ClearFooters left all three
' zones untouched and raised nothing. Twelve property writes cost a second or
' two. Pay it.
Public Function PanelPageSetup(ByVal ws As Worksheet, _
                               ByVal obj As ChartObject) As String
    Dim area As String
    area = ws.Range(obj.TopLeftCell, obj.BottomRightCell).Address

    With ws.PageSetup
        .PrintArea = area
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = 18                 ' 0.25 inch
        .RightMargin = 18
        .TopMargin = 18
        .BottomMargin = 18
        .HeaderMargin = 9
        .FooterMargin = 9
        .CenterHorizontally = True
        .CenterVertically = True
        .PrintGridlines = False
    End With

    PanelPageSetup = area
End Function


' ============================================================================
'  Verify - measure what Excel actually did
'
'  A chart that looks plausible and a chart that is right are different claims.
'  Returns "" or a newline-joined list of problems.
' ============================================================================

Public Function PanelVerify(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                            ByRef ly As TPanelLayout, ByVal ch As Chart) As String

    Dim problems As String
    Dim want As Long, got As Long, i As Long, n As Long
    Dim ser As Series
    Dim secCat As Axis, secVal As Axis, va As Axis
    Dim floorVal As Double

    want = PS_NSeries(sp) + mAnnotationSeries
    If PS_StacksElements(sp) Then want = want + 1
    got = ch.SeriesCollection.Count
    If got <> want Then
        problems = Add(problems, "series count " & got & ", expected " & want)
    End If

    ' No series may exceed the category count: past that Excel silently renders
    ' only the first NSlots points.
    For i = 1 To got
        Set ser = ch.SeriesCollection(i)
        n = -1
        On Error Resume Next
        n = ser.Points.Count
        Err.Clear
        On Error GoTo 0
        If n > PS_NSlots(sp) Then
            problems = Add(problems, "series " & i & " ('" & ser.name & "') has " & _
                n & " points against " & PS_NSlots(sp) & " categories - Excel " & _
                "will draw only the first " & PS_NSlots(sp))
        End If
    Next i

    If PS_Transposed(sp) Then
        Set secCat = ch.Axes(xlValue, xlSecondary)
        Set secVal = ch.Axes(xlCategory, xlSecondary)
    Else
        Set secCat = ch.Axes(xlCategory, xlSecondary)
        Set secVal = ch.Axes(xlValue, xlSecondary)
    End If

    If Abs(secCat.MinimumScale - 0.5) > 0.000000001 Or _
       Abs(secCat.MaximumScale - (PS_NSlots(sp) + 0.5)) > 0.000000001 Then
        problems = Add(problems, "secondary category axis is " & _
            secCat.MinimumScale & ".." & secCat.MaximumScale & ", must be 0.5.." & _
            (PS_NSlots(sp) + 0.5) & " or every annotation is off its category")
    End If

    floorVal = -PS_FloorPad(sp)
    Set va = ch.Axes(xlValue, xlPrimary)

    If Abs(va.MinimumScale - floorVal) > 0.000000001 Or _
       Abs(va.MaximumScale - PS_Bands(sp)) > 0.000000001 Then
        problems = Add(problems, "primary value axis is " & va.MinimumScale & _
            ".." & va.MaximumScale & ", must be " & floorVal & ".." & PS_Bands(sp))
    End If

    If Abs(secVal.MinimumScale - floorVal) > 0.000000001 Or _
       Abs(secVal.MaximumScale - PS_Bands(sp)) > 0.000000001 Then
        problems = Add(problems, "secondary value axis is " & secVal.MinimumScale & _
            ".." & secVal.MaximumScale & ", must match the primary or every " & _
            "annotation drifts off its data")
    End If

    ' The plot table may hold #N/A and nothing else.
    '
    ' VBA improves on the Python here. A #N/A read into a Variant is an Error
    ' subtype, so this is an EXACT test - the Python had to use a "v < -1e6"
    ' magic-number heuristic because it was reading through COM into ints.
    Dim vals As Variant, r As Long, c As Long, bad As Long
    Dim naText As String
    naText = CStr(CVErr(xlErrNA))
    vals = ws.Range(ws.Cells(ly.ptR0, ly.vc0), _
                    ws.Cells(ly.ptR1, ly.anC0 - 2)).value
    If IsArray(vals) Then
        For r = LBound(vals, 1) To UBound(vals, 1)
            For c = LBound(vals, 2) To UBound(vals, 2)
                If IsError(vals(r, c)) Then
                    If CStr(vals(r, c)) <> naText Then bad = bad + 1
                End If
            Next c
        Next r
    End If
    If bad > 0 Then
        problems = Add(problems, bad & " error cells in the plot table that " & _
                                 "are not #N/A")
    End If

    PanelVerify = problems
End Function


' ============================================================================
'  Liveness - the one claim the whole design rests on
'
'  Paste your numbers over the input block and the chart follows. PanelVerify
'  cannot see this: series counts, axis scales and stray error cells are all
'  equally true of a workbook whose formulas have been frozen into values.
'
'  Drive one input well past the current maximum, require the shared scale to
'  GROW and a plotted value to MOVE, then restore and require the chart to come
'  home - derived and merely-responsive are different claims. Saves nothing,
'  and restores even on error.
'
'      Immediate window:   ?modPanelChart.PanelLiveness(ActiveSheet)
'
'  MAKE IT FAIL BEFORE BELIEVING IT: run it against a copy whose
'  UsedRange.Value = UsedRange.Value - formulas frozen, everything still
'  looking correct - and it must report that nothing moved. A check that passes
'  everything passes nothing.
' ============================================================================

Public Function PanelLiveness(ByVal ws As Worksheet) As String

    Dim probe As Range, vmaxCell As Range, watch As Range
    Dim before As Variant, scaleBefore As Double, watchBefore As Variant
    Dim scaleAfter As Double, watchAfter As Variant
    Dim restored As Boolean
    Dim problem As String

    Set vmaxCell = FindLabelled(ws, "vmax")
    If vmaxCell Is Nothing Then
        PanelLiveness = "could not find the 'vmax' parameter - is this a panel sheet?"
        Exit Function
    End If

    ' Find the input block by its label rather than by address: a grid built
    ' with an origin shift puts it wherever it was told to.
    Set probe = FindInputProbe(ws)
    If probe Is Nothing Then
        PanelLiveness = "could not find the input block"
        Exit Function
    End If

    ' A plotted cell that should move when the scale does.
    Set watch = FindFirstPlottedValue(ws)
    If watch Is Nothing Then
        PanelLiveness = "could not find a plotted value column"
        Exit Function
    End If

    On Error GoTo Restore
    before = probe.value
    scaleBefore = CDbl(vmaxCell.value)
    watchBefore = watch.value

    probe.value = scaleBefore * 10 + 1000
    ws.Calculate
    scaleAfter = CDbl(vmaxCell.value)
    watchAfter = watch.value

    If scaleAfter <= scaleBefore Then
        problem = "the shared scale did not grow when the input block changed " & _
                  "- the engine is not reading the input block"
    ElseIf IsError(watchBefore) Or IsError(watchAfter) Then
        ' A gap either side is fine; both being #N/A tells us nothing.
        If IsError(watchBefore) And IsError(watchAfter) Then
            problem = "the watched value column is #N/A before and after - " & _
                      "cannot tell whether it moved"
        End If
    ElseIf Abs(CDbl(watchAfter) - CDbl(watchBefore)) < 0.000000001 Then
        problem = "no plotted series moved after the input changed - the " & _
                  "formulas look live but are not"
    End If

Restore:
    Dim errNum As Long, errTxt As String
    errNum = Err.Number
    errTxt = Err.description
    On Error Resume Next
    probe.value = before
    ws.Calculate
    restored = (Abs(CDbl(vmaxCell.value) - scaleBefore) < 0.000000001)
    Err.Clear
    On Error GoTo 0

    If errNum <> 0 Then
        PanelLiveness = "liveness check failed: " & errTxt
    ElseIf Len(problem) > 0 Then
        PanelLiveness = problem
    ElseIf Not restored Then
        PanelLiveness = "the input was restored but the scale did not come " & _
                        "back - check " & probe.Address & " by hand"
    Else
        PanelLiveness = ""
    End If
End Function


Private Function FindLabelled(ByVal ws As Worksheet, _
                              ByVal label As String) As Range
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:=label, LookIn:=xlValues, LookAt:=xlWhole, _
                          MatchCase:=False)
    Err.Clear
    On Error GoTo 0
    If Not f Is Nothing Then Set FindLabelled = f.offset(0, 1)
End Function


' The first numeric cell of the input block: one row below the "Panel" header
' and one column right of it.
Private Function FindInputProbe(ByVal ws As Worksheet) As Range
    Dim hdr As Range, c As Range
    On Error Resume Next
    Set hdr = ws.Cells.Find(What:="Panel", LookIn:=xlValues, LookAt:=xlWhole, _
                            MatchCase:=True)
    Err.Clear
    On Error GoTo 0
    If hdr Is Nothing Then Exit Function

    Set c = hdr.offset(1, 1)
    If IsNumeric(c.value) And Not IsEmpty(c.value) Then Set FindInputProbe = c
End Function


' The first cell of the first band's value column, found from the plot table's
' bold header row rather than by address.
Private Function FindFirstPlottedValue(ByVal ws As Worksheet) As Range
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="b0 e0", LookIn:=xlValues, LookAt:=xlPart, _
                          MatchCase:=False)
    Err.Clear
    On Error GoTo 0
    If f Is Nothing Then Exit Function

    ' Walk down to the first row that is not a gap - a separator slot is #N/A
    ' by design and would tell us nothing.
    Dim i As Long
    For i = 1 To 200
        If Not IsError(f.offset(i, 0).value) Then
            If IsNumeric(f.offset(i, 0).value) And Not IsEmpty(f.offset(i, 0).value) Then
                Set FindFirstPlottedValue = f.offset(i, 0)
                Exit Function
            End If
        End If
    Next i
End Function


' ============================================================================
'  Small helpers
' ============================================================================

Private Function Add(ByVal soFar As String, ByVal msg As String) As String
    If Len(soFar) = 0 Then
        Add = msg
    Else
        Add = soFar & vbCrLf & msg
    End If
End Function


' A panel with no name gets no title label - a label linked to a blank cell
' prints "0" where a name should be.
Private Function PanelNamed(ByRef sp As TPanelSpec, ByVal index As Long) As Boolean
    On Error GoTo NotNamed
    If index < 0 Then GoTo NotNamed
    If index > UBound(sp.panelNames) - LBound(sp.panelNames) Then GoTo NotNamed
    PanelNamed = (Len(Trim$(sp.panelNames(LBound(sp.panelNames) + index))) > 0)
    Exit Function
NotNamed:
    PanelNamed = False
End Function


' One value out of the input array. Empty means "no value here" - a grid may
' hold fewer panels than cells, and a blank read as zero draws a flat panel
' that is not there.
Private Function Datum(ByRef vals As Variant, ByRef sp As TPanelSpec, _
                       ByVal panel As Long, ByVal element As Long, _
                       ByVal period As Long) As Variant
    Datum = Empty
    On Error GoTo NoValue

    Dim r As Long, c As Long
    r = LBound(vals, 1) + panel
    c = LBound(vals, 2) + element * sp.Periods + period
    If r > UBound(vals, 1) Or c > UBound(vals, 2) Then Exit Function

    Dim v As Variant
    v = vals(r, c)
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(v) = 0 Then Exit Function
    End If
    If Not IsNumeric(v) Then Exit Function
    Datum = v
    Exit Function
NoValue:
    Datum = Empty
End Function


' How many series the annotation layer cost on the last build, for PanelVerify.
Public Function PanelAnnotationSeriesCount() As Long
    PanelAnnotationSeriesCount = mAnnotationSeries
End Function
