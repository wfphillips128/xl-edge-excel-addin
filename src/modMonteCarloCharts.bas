Attribute VB_Name = "modMonteCarloCharts"
' =============================================================================
' modMonteCarloCharts - draws the three Monte Carlo charts. No user interface.
'
' The numbers never live here. Each chart reads a block spilled by one of the
' library's chart-data LAMBDAs:
'
'   fx.RiskChartHist     ONE spill for both histogram charts: the title, the
'                        P10 / P50 / P90 lines, then one row per bin (center,
'                        edges, in-band, tail, count, cum %, right-edge X)
'   fx.RiskTornado       the title, then one row per input, largest effect first
'
' TITLES ARE LINKED. Both blocks start with a title row (the tornado's subtitle
' sits beside its title), and the chart's title is a formula pointing at that
' cell. Rename the output's header cell and the title follows. A linked title
' takes one font for all of it, so the tornado's smaller subtitle is a separate
' text box, linked the same way.
'
' THE HISTOGRAM FOLLOWS ITS BLOCK. In fx.RiskChartHist everything above the bins
' (title, P-lines, headers - MC_CHART_HEAD_ROWS rows) sits at fixed cells, so
' the P-lines and their labels point at ordinary ranges. The bin columns grow or
' shrink with the formula's Bins, so those series point at hidden sheet-level
' names instead - McChN_Cat, McChN_Band, ... - each one column of the spill:
'     =INDEX(Sheet!$J$2#, 7, 4):INDEX(Sheet!$J$2#, ROWS(Sheet!$J$2#), 4)
' Change Bins and the chart redraws with the new bins, nothing to re-insert.
' The tornado keeps plain ranges: its size is fixed by the inputs picked.
'
' THE ALIGNMENT CONTRACT. Percentile lines and the S-curve are XY scatter series
' on the SECONDARY axes, drawn over a column chart. With the secondary x axis
' pinned to 0 .. 1, bin i spans (i - 1) / B .. i / B of it for any bin count B,
' so a value v belongs at (v - left edge) / (B * width) - which is what the
' LAMBDA returns. The limits never depend on the data or on B, so they stay
' right after every recalculation. Do not "tidy" them to fit the data.
'
' ORDERING. Axes(xlCategory, xlSecondary) does not exist until a series has
' AxisGroup = 2, so the scatter series go on before any secondary axis is
' touched. And a secondary axis must be HIDDEN, never deleted: delete the
' secondary x axis and Excel quietly replots the scatter on the primary
' category axis, which slides every line out of place.
' =============================================================================
Option Explicit

' XL Edge brand, on a white background. Longs because VBA constants cannot
' call RGB(). The comment beside each is the color it encodes.
Private Const C_ACCENT As Long = 2843348       ' RGB(212,98,43)   the orange
Private Const C_ACTION As Long = 1854128       ' RGB(176,74,28)   darker orange, for text
Private Const C_INK As Long = 657930           ' RGB(10,10,10)    near-black
Private Const C_MUTED As Long = 5395026        ' RGB(82,82,82)    axis text
Private Const C_NEUTRAL As Long = 9211020      ' RGB(140,140,140) tail bars
Private Const C_RULE As Long = 15066597        ' RGB(229,229,229) gridlines
Private Const C_WHITE As Long = 16777215

Private Const CHART_W As Double = 560
Private Const CHART_H As Double = 340
Private Const TITLE_PT As Single = 16
Private Const SUBTITLE_PT As Single = 10.5

' Columns of the fx.RiskChartHist bin rows (1-based, as INDEX counts them).
Private Const COL_CENTER As Long = 1
Private Const COL_BAND As Long = 4
Private Const COL_TAIL As Long = 5
Private Const COL_COUNT As Long = 6
Private Const COL_CUM As Long = 7
Private Const COL_EDGE_X As Long = 8


' --- the three charts --------------------------------------------------------

' Histogram of outcomes with percentile lines. block is the first cell of the
' fx.RiskChartHist spill. Bars whose center lies between the lowest and highest
' percentile are orange, the tails gray; the band's edges sit on the lines.
Public Function McDrawOutcomeHistogram(ByVal ws As Worksheet, ByVal block As Range, _
                                       ByVal nMarks As Long, _
                                       ByVal atLeft As Double, ByVal atTop As Double) As ChartObject
    Dim obj As ChartObject, ch As Chart, ser As Series
    Dim tag As String, mark As Range
    Dim i As Long

    Set obj = NewChart(ws, atLeft, atTop, xlColumnClustered)
    Set ch = obj.Chart
    tag = NewTag(ws, obj)

    AddColumns ch, BinColumn(ws, block, tag, "Cat", COL_CENTER), _
               BinColumn(ws, block, tag, "Band", COL_BAND), "P10 to P90", C_ACCENT
    AddColumns ch, BinColumn(ws, block, tag, "Cat", COL_CENTER), _
               BinColumn(ws, block, tag, "Tail", COL_TAIL), "Tails", C_NEUTRAL
    ch.ChartGroups(1).Overlap = 100
    ch.ChartGroups(1).GapWidth = 8

    ' Every percentile line is the same black dashed line with a black label.
    ' The P-line rows sit at fixed cells, just under the title and a header.
    ' The outer-left label sits left of its line.
    For i = 1 To nMarks
        Set mark = block.offset(1 + i, 0)
        Set ser = AddLine(ch, mark.offset(0, 2).Resize(1, 2), mark.offset(0, 4).Resize(1, 2), _
                          CStr(mark.Value2), C_INK, msoLineDash)
        LabelTopPoint ser, mark, C_INK, _
                      IIf(i = 1 And nMarks > 1, xlLabelPositionLeft, xlLabelPositionRight)
    Next i

    PinSecondary ch, 0, 1, False
    StyleCategoryAxis ch, block
    StyleValueAxis ch.Axes(xlValue, xlPrimary)
    LinkTitle ch, block
    ch.HasLegend = False
    Set McDrawOutcomeHistogram = obj
End Function

' Histogram with the S-curve (cumulative %) on a right-hand axis, both from the
' same fx.RiskChartHist block. The curve runs through each bin's right edge at
' the share of trials up to that edge - exact, and always on the bars' grid.
Public Function McDrawHistogramSCurve(ByVal ws As Worksheet, ByVal block As Range, _
                                      ByVal atLeft As Double, ByVal atTop As Double) As ChartObject
    Dim obj As ChartObject, ch As Chart, ser As Series
    Dim tag As String

    Set obj = NewChart(ws, atLeft, atTop, xlColumnClustered)
    Set ch = obj.Chart
    tag = NewTag(ws, obj)

    AddColumns ch, BinColumn(ws, block, tag, "Cat", COL_CENTER), _
               BinColumn(ws, block, tag, "Count", COL_COUNT), "Trials per bin", C_NEUTRAL
    ch.ChartGroups(1).GapWidth = 8

    Set ser = AddLine(ch, BinColumn(ws, block, tag, "X", COL_EDGE_X), _
                      BinColumn(ws, block, tag, "Cum", COL_CUM), "Cumulative %", C_ACCENT, msoLineSolid)
    ser.ChartType = xlXYScatterSmoothNoMarkers
    ser.Format.Line.weight = 2.5

    PinSecondary ch, 0, 1, True
    StyleCategoryAxis ch, block
    StyleValueAxis ch.Axes(xlValue, xlPrimary)
    LinkTitle ch, block
    ch.HasLegend = False
    Set McDrawHistogramSCurve = obj
End Function

' Tornado from an fx.RiskTornado block. titleCell is its first cell, with the
' subtitle beside it; topCell is its header cell; method is "swing" or "rank".
' The largest effect is drawn at the top.
Public Function McDrawTornado(ByVal ws As Worksheet, ByVal titleCell As Range, _
                              ByVal topCell As Range, ByVal nInputs As Long, _
                              ByVal method As String, _
                              ByVal atLeft As Double, ByVal atTop As Double) As ChartObject
    Dim obj As ChartObject, ch As Chart
    Dim names As Range, s1 As Series, s2 As Series
    Dim h As Double

    ' Height grows with the number of inputs so the bars stay readable.
    h = Application.WorksheetFunction.Max(CHART_H, 110 + 26 * nInputs)
    Set obj = NewChart(ws, atLeft, atTop, xlBarClustered, h)
    Set ch = obj.Chart
    Set names = topCell.offset(1, 0).Resize(nInputs, 1)

    If LCase$(method) = "rank" Then
        Set s1 = AddBars(ch, names, topCell.offset(1, 2).Resize(nInputs, 1), "Positive", C_ACCENT)
        Set s2 = AddBars(ch, names, topCell.offset(1, 3).Resize(nInputs, 1), "Negative", C_INK)
        LabelEachPoint s1, topCell.offset(1, 1), nInputs, True, False
        With ch.Axes(xlValue, xlPrimary)
            .MinimumScale = -1
            .MaximumScale = 1
            .TickLabels.NumberFormat = "0.0"
        End With
        ch.HasLegend = False
    Else
        ' Plotted as change from the base (the output's mean), so the axis
        ' crosses at 0 - a constant, which stays right when the model changes.
        ' The labels read the absolute Low/High means straight from the cells.
        Set s1 = AddBars(ch, names, topCell.offset(1, 3).Resize(nInputs, 1), _
                         "Input in its lowest 10%", C_NEUTRAL)
        Set s2 = AddBars(ch, names, topCell.offset(1, 4).Resize(nInputs, 1), _
                         "Input in its highest 10%", C_ACCENT)
        LabelEachPoint s1, topCell.offset(1, 1), nInputs, False, True
        LabelEachPoint s2, topCell.offset(1, 2), nInputs, False, True
        ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "+General;-General;0"
        ch.HasLegend = True
        ch.Legend.Position = xlLegendPositionBottom
        ch.Legend.Font.color = C_MUTED
    End If

    ch.ChartGroups(1).Overlap = 100
    ch.ChartGroups(1).GapWidth = 40
    With ch.Axes(xlCategory, xlPrimary)
        .ReversePlotOrder = True                   ' first row (largest) at the top
        .Crosses = xlMaximum                       ' ...which keeps the value axis at the bottom
        .TickLabelPosition = xlTickLabelPositionLow ' names at the left, not at 0
        .MajorTickMark = xlTickMarkNone
        .TickLabels.Font.color = C_INK
        .TickLabels.Font.Size = 10
    End With
    StyleValueAxis ch.Axes(xlValue, xlPrimary)
    LinkTitle ch, titleCell
    LinkSubtitle ch, titleCell.offset(0, 1)
    Set McDrawTornado = obj
End Function


' --- building blocks ----------------------------------------------------------

Private Function NewChart(ByVal ws As Worksheet, ByVal atLeft As Double, ByVal atTop As Double, _
                          ByVal kind As XlChartType, Optional ByVal h As Double = 0) As ChartObject
    Dim obj As ChartObject
    If h <= 0 Then h = CHART_H
    Set obj = ws.ChartObjects.Add(atLeft, atTop, CHART_W, h)
    With obj.Chart
        .ChartType = kind
        Do While .SeriesCollection.count > 0          ' Excel may guess a series from the selection
            .SeriesCollection(1).Delete
        Loop
        .ChartArea.Format.Fill.ForeColor.RGB = C_WHITE
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
        .ChartArea.Font.color = C_INK
    End With
    Set NewChart = obj
End Function

' A free tag for one chart's names - McCh1, McCh2, ... The chart object is
' named to match (McChart1, ...), so a name and its chart are easy to pair up.
Private Function NewTag(ByVal ws As Worksheet, ByVal obj As ChartObject) As String
    Dim n As Long
    Do
        n = n + 1
    Loop While SheetNameExists(ws, "McCh" & n & "_Cat") Or ChartNamed(ws, "McChart" & n)
    obj.name = "McChart" & n
    NewTag = "McCh" & n
End Function

' A series reference to one column of the bins in a fx.RiskChartHist spill,
' through a hidden sheet-level name that follows the spill as Bins changes.
' Creates the name the first time it is asked for. col counts from 1.
Private Function BinColumn(ByVal ws As Worksheet, ByVal block As Range, ByVal tag As String, _
                           ByVal key As String, ByVal col As Long) As String
    Dim sheetRef As String, spill As String, nm As String
    sheetRef = "'" & Replace(ws.name, "'", "''") & "'!"
    spill = sheetRef & block.Cells(1, 1).Address & "#"
    nm = tag & "_" & key
    If Not SheetNameExists(ws, nm) Then
        ws.names.Add name:=nm, Visible:=False, _
            RefersTo:="=INDEX(" & spill & "," & (modMonteCarlo.MC_CHART_HEAD_ROWS + 1) & "," & col & _
                      "):INDEX(" & spill & ",ROWS(" & spill & ")," & col & ")"
    End If
    BinColumn = "=" & sheetRef & nm
End Function

' Sheet-level names report themselves as "Sheet!name", so match on the part
' after the "!".
Private Function SheetNameExists(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    Dim n As name
    For Each n In ws.names
        If StrComp(Mid$(n.name, InStrRev(n.name, "!") + 1), nm, vbTextCompare) = 0 Then
            SheetNameExists = True
            Exit Function
        End If
    Next n
End Function

Private Function ChartNamed(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    Dim obj As ChartObject
    For Each obj In ws.ChartObjects
        If StrComp(obj.name, nm, vbTextCompare) = 0 Then
            ChartNamed = True
            Exit Function
        End If
    Next obj
End Function

' One column series. xRef and yRef are series references from BinColumn.
Private Sub AddColumns(ByVal ch As Chart, ByVal xRef As String, ByVal yRef As String, _
                       ByVal nm As String, ByVal color As Long)
    Dim ser As Series
    Set ser = ch.SeriesCollection.NewSeries
    ser.name = nm
    ser.XValues = xRef
    ser.Values = yRef
    ser.Format.Fill.ForeColor.RGB = color
    ser.Format.Line.Visible = msoFalse
End Sub

Private Function AddBars(ByVal ch As Chart, ByVal names As Range, ByVal vals As Range, _
                         ByVal nm As String, ByVal color As Long) As Series
    Dim ser As Series
    Set ser = ch.SeriesCollection.NewSeries
    ser.name = nm
    ser.XValues = names
    ser.Values = vals
    ser.Format.Fill.ForeColor.RGB = color
    ser.Format.Line.Visible = msoFalse
    Set AddBars = ser
End Function

' A scatter line on the secondary axes. ChartType before AxisGroup, and
' AxisGroup before anything touches a secondary axis. xs and ys are ranges, or
' series references from BinColumn. A range must reach the series through a
' Range variable: handed over inside a Variant it raises Type mismatch.
Private Function AddLine(ByVal ch As Chart, ByVal xs As Variant, ByVal ys As Variant, _
                         ByVal nm As String, ByVal color As Long, _
                         ByVal dash As MsoLineDashStyle) As Series
    Dim ser As Series, rg As Range
    Set ser = ch.SeriesCollection.NewSeries
    ser.ChartType = xlXYScatterLinesNoMarkers
    ser.AxisGroup = xlSecondary
    ser.name = nm
    If IsObject(xs) Then
        Set rg = xs
        ser.XValues = rg
        Set rg = ys
        ser.Values = rg
    Else
        ser.XValues = xs
        ser.Values = ys
    End If
    With ser.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = color
        .weight = 1.75
        .DashStyle = dash
    End With
    Set AddLine = ser
End Function

' The label on a percentile line is LINKED to its Label cell ("P50  103.9"),
' so it updates with the model like everything else.
Private Sub LabelTopPoint(ByVal ser As Series, ByVal labelCell As Range, _
                          ByVal color As Long, ByVal pos As XlDataLabelPosition)
    Dim pt As Point
    Set pt = ser.Points(2)
    pt.HasDataLabel = True
    pt.DataLabel.formula = "=" & labelCell.Address(External:=True)
    pt.DataLabel.Position = pos
    pt.DataLabel.Font.color = color
    pt.DataLabel.Font.bold = True
    pt.DataLabel.Font.Size = 10
End Sub

' Data label on every bar, linked to the matching cell in a column of the block.
' Inside the bar end, in white, where a label outside the end would collide
' with the input names at the plot's left edge.
Private Sub LabelEachPoint(ByVal ser As Series, ByVal firstCell As Range, ByVal n As Long, _
                           ByVal asCorrelation As Boolean, ByVal inside As Boolean)
    Dim i As Long, pt As Point
    For i = 1 To n
        Set pt = ser.Points(i)
        pt.HasDataLabel = True
        pt.DataLabel.formula = "=" & firstCell.offset(i - 1, 0).Address(External:=True)
        pt.DataLabel.Font.Size = 9
        If inside Then
            pt.DataLabel.Position = xlLabelPositionInsideEnd
            pt.DataLabel.Font.color = C_WHITE
            pt.DataLabel.Font.bold = True
        Else
            pt.DataLabel.Position = xlLabelPositionOutsideEnd
            pt.DataLabel.Font.color = C_MUTED
        End If
    Next i
    ' A linked label shows the cell's own formatting, so give the source a
    ' sensible one. This formats cells inside our own spill, nothing else.
    firstCell.Resize(n, 1).NumberFormat = IIf(asCorrelation, "0.00", "#,##0.0")
End Sub

' THE ALIGNMENT CONTRACT - see the module header.
Private Sub PinSecondary(ByVal ch As Chart, ByVal yMin As Double, _
                         ByVal yMax As Double, ByVal showPercentAxis As Boolean)
    Dim sx As Axis, sy As Axis
    ch.HasAxis(xlCategory, xlSecondary) = True
    ch.HasAxis(xlValue, xlSecondary) = True
    Set sx = ch.Axes(xlCategory, xlSecondary)
    Set sy = ch.Axes(xlValue, xlSecondary)

    sx.MinimumScale = 0
    sx.MaximumScale = 1
    sx.TickLabelPosition = xlTickLabelPositionNone
    sx.MajorTickMark = xlTickMarkNone
    HideLine sx

    sy.MinimumScale = yMin
    sy.MaximumScale = yMax
    If showPercentAxis Then
        sy.TickLabels.NumberFormat = "0%"
        sy.TickLabels.Font.color = C_ACTION
        sy.MajorUnit = 0.25
        sy.MajorTickMark = xlTickMarkNone
    Else
        sy.TickLabelPosition = xlTickLabelPositionNone
        sy.MajorTickMark = xlTickMarkNone
    End If
    HideLine sy
End Sub

' Bin centers along the bottom: one label every fifth bin, formatted to the
' size of the numbers as they are now - judged from the first bin and the
' upper P-line, which sit at fixed cells of the block.
Private Sub StyleCategoryAxis(ByVal ch As Chart, ByVal block As Range)
    Dim ax As Axis, biggest As Double
    biggest = Application.WorksheetFunction.Max( _
                  Abs(block.offset(modMonteCarlo.MC_CHART_HEAD_ROWS, 0).Value2), _
                  Abs(block.offset(1 + modMonteCarlo.MC_CHART_MARKS, 1).Value2))
    Set ax = ch.Axes(xlCategory, xlPrimary)
    ax.TickLabelSpacing = 5
    ax.TickMarkSpacing = 5
    ax.MajorTickMark = xlTickMarkOutside
    ax.TickLabels.NumberFormat = IIf(biggest >= 100, "#,##0", IIf(biggest >= 1, "#,##0.0", "0.000"))
    ax.TickLabels.Font.color = C_MUTED
    ax.TickLabels.Font.Size = 9
    ax.Format.Line.ForeColor.RGB = C_INK
End Sub

Private Sub StyleValueAxis(ByVal ax As Axis)
    ax.TickLabels.Font.color = C_MUTED
    ax.TickLabels.Font.Size = 9
    ax.MajorTickMark = xlTickMarkNone
    HideLine ax
    If ax.HasMajorGridlines Then
        ax.MajorGridlines.Format.Line.ForeColor.RGB = C_RULE
    End If
End Sub

' Title in bold ink, LINKED to titleCell, so it follows the cell.
' Format it through ChartTitle.Font only: touching ChartTitle.Format.TextFrame2
' silently turns a linked title back into fixed text.
Private Sub LinkTitle(ByVal ch As Chart, ByVal titleCell As Range)
    ch.HasTitle = True
    ch.chartTitle.formula = "=" & titleCell.Address(External:=True)
    With ch.chartTitle.Font
        .color = C_INK
        .Size = TITLE_PT
        .bold = True
    End With
    ch.chartTitle.HorizontalAlignment = xlLeft
    ch.chartTitle.Left = 8
End Sub

' A smaller, muted line under the title, in a text box linked to subtitleCell.
' The plot area is moved down by the box's height so nothing overlaps.
Private Sub LinkSubtitle(ByVal ch As Chart, ByVal subtitleCell As Range)
    Dim shp As Shape
    Dim boxTop As Double, newTop As Double

    boxTop = ch.chartTitle.Top + ch.chartTitle.Height - 2
    Set shp = ch.shapes.AddTextbox(msoTextOrientationHorizontal, ch.chartTitle.Left + 2, boxTop, _
                                   ch.ChartArea.width - ch.chartTitle.Left - 18, SUBTITLE_PT * 2)
    shp.name = "McSubtitle"
    shp.DrawingObject.formula = "=" & subtitleCell.Address(External:=True)
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    With shp.TextFrame2
        .WordWrap = msoTrue
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        With .TextRange.Font
            .Size = SUBTITLE_PT
            .bold = msoFalse
            .Fill.ForeColor.RGB = C_MUTED
        End With
    End With

    newTop = boxTop + shp.Height + 6
    If ch.PlotArea.Top < newTop Then
        ch.PlotArea.Height = ch.PlotArea.Height - (newTop - ch.PlotArea.Top)
        ch.PlotArea.Top = newTop
    End If
End Sub

' Hide an axis line whichever way this Excel build allows. As Object because
' neither Axis.Format nor Axis.Border is reachable on the secondary x axis in
' every build (both can raise E_FAIL), and a cosmetic line is not worth failing
' a chart over. Same approach as modPanelChart.
Private Sub HideLine(ByVal o As Object)
    On Error Resume Next
    o.Format.Line.Visible = msoFalse
    If Err.Number <> 0 Then
        Err.Clear
        o.Border.LineStyle = xlNone
    End If
    Err.Clear
    On Error GoTo 0
End Sub
