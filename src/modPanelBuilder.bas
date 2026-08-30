Attribute VB_Name = "modPanelBuilder"
Option Explicit

' ============================================================================
'  modPanelBuilder
'  Writes one panel-chart spec into one worksheet: the input block the reader
'  edits, the engine that fans it out, and the annotation tables the chart
'  draws its furniture from. Everything except the chart object itself, which
'  is modPanelChart's job.
'
'  ONE WRITE PER COLUMN, NEVER PER CELL
'  ------------------------------------
'  The plot table for a seven-row grid is around 13,600 cells. Written cell by
'  cell that is 13,600 COM round trips and the tool gets abandoned; written
'  column by column it is about forty. Every row of a column has the same
'  formula SHAPE, so in R1C1 it is the same string and the whole column is one
'  scalar assignment - see modPanelFormulas for why that mode exists and what
'  it costs.
'
'  A 5x4x12x2 line grid is roughly 250 COM calls end to end, of which about 130
'  are per-series chart property sets that cannot be batched.
'
'  Ported from panel_excel.py (PanelBuilder._write_*) in the panel-charts
'  skill. Keep the two in step.
' ============================================================================

' An annotation family, as handed to the chart layer. The chart needs to know
' where the points are and how many rows must stay together when it splits the
' family across several series - see the render cap in modPanelChart.
Public Type TAnnotation
    name    As String
    col     As Long          ' x column; y is col + 1
    labCol  As Long          ' label text column, or 0
    r0      As Long
    n       As Long
    unit    As Long          ' rows that must stay together: 3 polyline, 1 point
End Type

Public Const AN_MAX As Long = 6

' Which formula a column carries. VBA has no function pointers, so the write
' helper dispatches on this rather than taking a composer - which keeps the
' single-assignment fast path and the per-row trace path in ONE place instead
' of repeating the If PanelAddressMode() branch at every column.
Private Enum eFx
    efPos = 1
    efBlk = 2
    efSub = 3
    efK = 4
    efE = 5
    efSep = 6
    efKx = 7
    efBlkx = 8
    efCat = 9
    efValue = 10
    efBase = 11
    efMask = 12
    efPinValue = 13
    efPinStem = 14
    efPointLabel = 15
End Enum

' Set once per build by the AGGREGATE probe; read by the stacked base column.
Private mAggregateOK As Boolean


' ============================================================================
'  Entry point
' ============================================================================

' Write the whole engine into `ws`. Returns "" or a problem.
'
' `vals` is a 1-based two-dimensional Variant, n_panels rows by
' elements*periods columns. Pass Empty for placeholder data.
'
' `originShift` pushes every row the engine writes down by that many, so a
' caller who already owns the top of the sheet can host the grid underneath.
' Only rows move; the engine's columns are its own either way.
Public Function PanelWriteSheet(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                                ByRef st As TPanelStyle, ByRef vals As Variant, _
                                ByRef ly As TPanelLayout, _
                                ByRef anns() As TAnnotation, _
                                ByRef anCount As Long, _
                                Optional ByVal originShift As Long = 0, _
                                Optional ByVal intro As Boolean = True) As String

    Dim problem As String

    problem = PanelSpecInit(sp)
    If Len(problem) > 0 Then PanelWriteSheet = problem: Exit Function

    problem = LayoutInit(sp, ly, originShift)
    If Len(problem) > 0 Then PanelWriteSheet = problem: Exit Function

    If IsEmpty(vals) Then vals = PanelDummyData(sp)

    mAggregateOK = ProbeAggregate(ws, ly)

    If intro Then WriteIntro ws, sp, ly
    WriteInput ws, sp, ly, vals
    If PS_StacksElements(sp) Then WriteStackSums ws, sp, ly
    WriteParams ws, sp, ly
    WriteHelperBlock ws, sp, st, ly
    WriteValueBlock ws, sp, st, ly
    WriteAnnotations ws, sp, st, ly, anns, anCount

    ' FastModeOn left calculation manual, so nothing written above has a value
    ' yet. Everything the chart layer does next reads VALUES - axis bounds, the
    ' stray-error scan, the AGGREGATE result - so the sheet has to be brought
    ' up to date first. The Python engine has nothing to say here because it
    ' runs with calculation automatic throughout.
    ws.Calculate

    PanelWriteSheet = ""
End Function


' ============================================================================
'  The input block - the only cells the reader ever edits
' ============================================================================

Private Sub WriteIntro(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                       ByRef ly As TPanelLayout)
    Dim t As String
    t = sp.title
    If Len(t) = 0 Then
        t = KindLabel(sp.kind) & " panel chart - " & sp.rows & " x " & sp.cols
    End If

    With ws.Cells(1 + ly.originShift, 1)
        .value = t
        .Font.bold = True
        .Font.Size = 14
    End With
    ws.Cells(2 + ly.originShift, 1).value = _
        "Paste your own numbers over the block below. Everything else is " & _
        "worked out from it - do not edit the engine columns to the right."
End Sub


Private Sub WriteInput(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                       ByRef ly As TPanelLayout, ByRef vals As Variant)

    Dim e As Long, k As Long, c0 As Long
    Dim arr() As Variant

    With ws.Cells(ly.hdrPerRow, 1)
        .value = "Panel"
        .Font.bold = True
    End With

    For e = 0 To sp.Elements - 1
        c0 = ly.inC0 + e * sp.Periods

        With ws.Cells(ly.hdrElemRow, c0)
            .NumberFormat = "@"
            .value = sp.elementNames(LBound(sp.elementNames) + e)
            .Font.bold = True
        End With
        If sp.Periods > 1 Then
            ' Only the upper-left cell has content, so Merge does not prompt.
            ' DisplayAlerts is left alone: AppStateManager already captured and
            ' cleared it, and setting it back here would fight the restore.
            ws.Range(ws.Cells(ly.hdrElemRow, c0), _
                     ws.Cells(ly.hdrElemRow, c0 + sp.Periods - 1)).Merge
            ws.Cells(ly.hdrElemRow, c0).HorizontalAlignment = xlCenter
        End If

        ReDim arr(1 To 1, 1 To sp.Periods)
        For k = 0 To sp.Periods - 1
            arr(1, k + 1) = sp.periodLabels(LBound(sp.periodLabels) + k)
        Next k

        With ws.Range(ws.Cells(ly.hdrPerRow, c0), _
                      ws.Cells(ly.hdrPerRow, c0 + sp.Periods - 1))
            ' Text format BEFORE the write, or Excel eats the label. Anything
            ' date-shaped - "Jan-24", "Q1 24", "1/24", the ordinary way a
            ' finance period is written - is silently parsed into a date
            ' serial, and because the tick labels are LINKED TO THESE CELLS the
            ' chart then prints 46046 where the reader typed Jan-24. Nothing
            ' errors; the number just appears along the axis. Formatting
            ' afterwards does not undo it - the ordering IS the fix.
            .NumberFormat = "@"
            .value = arr
        End With
    Next e

    ' Same trap: a panel legitimately called "2024" or "Mar" is a label, not a
    ' quantity, and its title is linked to the cell.
    ReDim arr(1 To PS_NPanels(sp), 1 To 1)
    For k = 0 To PS_NPanels(sp) - 1
        arr(k + 1, 1) = sp.panelNames(LBound(sp.panelNames) + k)
    Next k
    With ws.Range(ws.Cells(ly.inR0, 1), ws.Cells(ly.inR1, 1))
        .NumberFormat = "@"
        .value = arr
        .ColumnWidth = 18
    End With

    With ws.Range(ws.Cells(ly.inR0, ly.inC0), ws.Cells(ly.inR1, ly.inC1))
        .value = vals
        .NumberFormat = "#,##0"
    End With

    ws.Range(ws.Cells(ly.hdrPerRow, 1), _
             ws.Cells(ly.hdrPerRow, ly.inC1)).Font.bold = True
End Sub


' Area only. Elements pile on each other, so the shared scale has to be sized
' on the stacked TOTAL rather than on the largest individual value - otherwise
' the top of every stack runs out of its band.
Private Sub WriteStackSums(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                           ByRef ly As TPanelLayout)

    Dim arr() As Variant
    Dim p As Long, k As Long, e As Long
    Dim s As String, Ref As String

    ws.Cells(ly.ssR0 - 1, ly.ssC0).value = "Stacked totals - the shared scale reads these"
    ws.Cells(ly.ssR0 - 1, ly.ssC0).Font.bold = True

    ReDim arr(1 To PS_NPanels(sp), 1 To sp.Periods)
    For p = 0 To PS_NPanels(sp) - 1
        For k = 0 To sp.Periods - 1
            s = ""
            For e = 0 To sp.Elements - 1
                Ref = ColLetters(ly.inC0 + e * sp.Periods + k) & (ly.inR0 + p)
                If Len(s) > 0 Then s = s & "+"
                s = s & Ref
            Next e
            arr(p + 1, k + 1) = "=" & s
        Next k
    Next p

    With ws.Range(ws.Cells(ly.ssR0, ly.ssC0), _
                  ws.Cells(ly.ssR1, ly.ssC0 + sp.Periods - 1))
        .formula = arr
        .NumberFormat = "#,##0"
    End With
End Sub


' ============================================================================
'  The engine parameters
' ============================================================================

Private Sub WriteParams(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                        ByRef ly As TPanelLayout)

    Dim labels() As Variant, vals() As Variant
    ReDim labels(1 To PARAM_ROWS, 1 To 1)
    ReDim vals(1 To PARAM_ROWS, 1 To 1)

    With ws.Cells(ly.parR0 - 1, ly.engC0)
        .value = "ENGINE - do not edit"
        .Font.bold = True
    End With

    labels(1, 1) = "grid rows":    vals(1, 1) = sp.rows
    labels(2, 1) = "grid cols":    vals(2, 1) = sp.cols
    labels(3, 1) = "periods":      vals(3, 1) = sp.Periods
    labels(4, 1) = "elements":     vals(4, 1) = sp.Elements
    labels(5, 1) = "blocks":       vals(5, 1) = PS_Blocks(sp)
    labels(6, 1) = "bands":        vals(6, 1) = PS_Bands(sp)
    labels(7, 1) = "sub-slots":    vals(7, 1) = PS_Subslots(sp)
    labels(8, 1) = "block slots":  vals(8, 1) = PS_BlockSlots(sp)
    labels(9, 1) = "spacers":      vals(9, 1) = PS_Spacers(sp)
    labels(10, 1) = "lead":        vals(10, 1) = sp.lead
    labels(11, 1) = "band frac":   vals(11, 1) = sp.BandFrac
    labels(12, 1) = "min raw":     vals(12, 1) = FxMinRaw(ly, sp)
    labels(13, 1) = "max raw":     vals(13, 1) = FxMaxRaw(ly, sp)
    labels(14, 1) = "step":        vals(14, 1) = FxStep(ly)
    labels(15, 1) = "vmin":        vals(15, 1) = FxVMin(ly, sp)
    labels(16, 1) = "vmax":        vals(16, 1) = FxVMax(ly)
    labels(17, 1) = "span":        vals(17, 1) = FxSpan(ly)
    labels(18, 1) = "zero offset": vals(18, 1) = FxZeroOffset(ly)

    ws.Range(ws.Cells(ly.parR0, ly.engC0), _
             ws.Cells(ly.parR0 + PARAM_ROWS - 1, ly.engC0)).value = labels

    ' The parameter values are written with .Formula, not .Value: rows 12-18
    ' are formula strings and the rest are numbers, and a mixed array assigned
    ' to .Formula keeps each as what it is.
    ws.Range(ws.Cells(ly.parR0, ly.parC0), _
             ws.Cells(ly.parR0 + PARAM_ROWS - 1, ly.parC0)).formula = vals

    ws.Cells(ly.parR0, ly.engC0).Resize(PARAM_ROWS, 1).Font.Italic = True
End Sub


' The probe.
'
' AGGREGATE is native from Excel 2010 and needs no _xlfn. prefix when a live
' Excel is doing the writing - that prefix exists for tools writing a file
' Excel has not yet parsed. The risk is not that it fails; it is that on an odd
' build it SUCCEEDS and stores something else. So write it, read it back, and
' fall through to the exact expansion if it did not survive.
Private Function ProbeAggregate(ByVal ws As Worksheet, _
                                ByRef ly As TPanelLayout) As Boolean
    Dim probe As Range
    Set probe = ws.Cells(1 + ly.originShift, ly.anC0 + 40)

    On Error GoTo NotAvailable
    probe.formula = "=AGGREGATE(9,6,1)"
    probe.Calculate

    ' Two ways this fails, and only the second is loud. If Excel rewrote the
    ' name, the formula reads back carrying the prefix; if it accepted a name
    ' it cannot resolve, the value is #NAME?.
    If InStr(1, probe.formula, "_xlfn.", vbTextCompare) > 0 Then GoTo NotAvailable
    If IsError(probe.Value2) Then GoTo NotAvailable
    If probe.Value2 <> 1 Then GoTo NotAvailable

    probe.ClearContents
    ProbeAggregate = True
    Exit Function

NotAvailable:
    On Error Resume Next
    probe.ClearContents
    Err.Clear
    ProbeAggregate = False
End Function


' ============================================================================
'  The plot table
' ============================================================================

Private Sub WriteHelperBlock(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                             ByRef st As TPanelStyle, ByRef ly As TPanelLayout)

    Dim n As Long, i As Long
    n = PS_NSlots(sp)

    WriteHeader ws, ly, ly.cX, "x"
    WriteHeader ws, ly, ly.cPos, "pos"
    WriteHeader ws, ly, ly.cBlk, "blk"
    WriteHeader ws, ly, ly.cSub, "sub"
    WriteHeader ws, ly, ly.cK, "k"
    WriteHeader ws, ly, ly.cE, "e"
    WriteHeader ws, ly, ly.cSep, "sep"
    WriteHeader ws, ly, ly.cKx, "kx"
    WriteHeader ws, ly, ly.cBlkx, "blkx"
    WriteHeader ws, ly, ly.cCat, "cat"

    ' The x column is the literal 1..n, written as VALUES. Not "=ROW()-ptR0+1":
    ' every other helper is an offset from this column, and a formula here
    ' would make the whole table depend on where it sits on the sheet.
    Dim xs() As Variant
    ReDim xs(1 To n, 1 To 1)
    For i = 1 To n
        xs(i, 1) = i
    Next i
    ws.Range(ws.Cells(ly.ptR0, ly.cX), ws.Cells(ly.ptR1, ly.cX)).Value2 = xs

    PutCol ws, sp, ly, ly.cPos, efPos, st
    PutCol ws, sp, ly, ly.cBlk, efBlk, st
    PutCol ws, sp, ly, ly.cSub, efSub, st
    PutCol ws, sp, ly, ly.cK, efK, st
    PutCol ws, sp, ly, ly.cE, efE, st
    PutCol ws, sp, ly, ly.cSep, efSep, st
    PutCol ws, sp, ly, ly.cKx, efKx, st
    PutCol ws, sp, ly, ly.cBlkx, efBlkx, st
    PutCol ws, sp, ly, ly.cCat, efCat, st
End Sub


Private Sub WriteValueBlock(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                            ByRef st As TPanelStyle, ByRef ly As TPanelLayout)

    Dim band As Long, e As Long, tone As Long, col As Long

    For band = 0 To PS_Bands(sp) - 1

        If PS_Stacked(sp) Then
            col = LY_BaseCol(ly, band)
            WriteHeader ws, ly, col, "base " & band
            PutCol ws, sp, ly, col, efBase, st, band
        End If

        For e = 0 To sp.Elements - 1
            For tone = 0 To ly.tones - 1
                col = LY_ValCol(ly, sp, band, e, tone)
                WriteHeader ws, ly, col, "b" & band & " e" & e & _
                            IIf(ly.tones > 1, IIf(tone = 0, " +", " -"), "")
                If PS_TwoTone(sp) Then
                    PutCol ws, sp, ly, col, efPinValue, st, band, e, tone
                Else
                    PutCol ws, sp, ly, col, efValue, st, band, e, tone
                End If
            Next tone
        Next e
    Next band

    If PS_StacksElements(sp) Then
        WriteHeader ws, ly, ly.cMask, "mask"
        PutCol ws, sp, ly, ly.cMask, efMask, st
    End If

    If PS_TwoTone(sp) Then
        For band = 0 To PS_Bands(sp) - 1
            For e = 0 To sp.Elements - 1
                For tone = 0 To ly.tones - 1
                    col = LY_ErrCol(ly, band, e, tone)
                    WriteHeader ws, ly, col, "stem b" & band & " e" & e
                    PutCol ws, sp, ly, col, efPinStem, st, band, e, tone
                Next tone
            Next e
        Next band
    End If

    If PS_Labelled(sp) Then
        For band = 0 To PS_Bands(sp) - 1
            For e = 0 To sp.Elements - 1
                For tone = 0 To ly.tones - 1
                    col = LY_LabCol(ly, band, e, tone)
                    WriteHeader ws, ly, col, "lab b" & band & " e" & e
                    PutCol ws, sp, ly, col, efPointLabel, st, band, e, tone
                Next tone
            Next e
        Next band
    End If
End Sub


Private Sub WriteHeader(ByVal ws As Worksheet, ByRef ly As TPanelLayout, _
                        ByVal col As Long, ByVal text As String)
    With ws.Cells(ly.ptR0 - 1, col)
        .value = text
        .Font.bold = True
        .Font.Size = 8
    End With
End Sub


' Write one whole plot-table column.
'
' In the shipping mode this is ONE assignment for the entire column, because
' every row carries the same R1C1 string. The trace mode builds the per-row A1
' array instead - same cells, diffable against the reference documents.
Private Sub PutCol(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                   ByRef ly As TPanelLayout, ByVal col As Long, _
                   ByVal which As eFx, _
                   ByRef st As TPanelStyle, _
                   Optional ByVal band As Long = 0, _
                   Optional ByVal element As Long = 0, _
                   Optional ByVal tone As Long = 0)

    Dim rg As Range
    Set rg = ws.Range(ws.Cells(ly.ptR0, col), ws.Cells(ly.ptR1, col))

    If PanelAddressMode() = amR1C1 Then
        rg.FormulaR1C1 = ComposeFx(which, sp, ly, st, band, element, tone, ly.ptR0)
        Exit Sub
    End If

    Dim n As Long, i As Long, v() As Variant
    n = ly.ptR1 - ly.ptR0 + 1
    ReDim v(1 To n, 1 To 1)
    For i = 1 To n
        v(i, 1) = ComposeFx(which, sp, ly, st, band, element, tone, ly.ptR0 + i - 1)
    Next i
    rg.formula = v
End Sub


Private Function ComposeFx(ByVal which As eFx, ByRef sp As TPanelSpec, _
                           ByRef ly As TPanelLayout, ByRef st As TPanelStyle, _
                           ByVal band As Long, ByVal element As Long, _
                           ByVal tone As Long, ByVal r As Long) As String
    Select Case which
        Case efPos:        ComposeFx = FxPos(ly, r)
        Case efBlk:        ComposeFx = FxBlk(ly, r)
        Case efSub:        ComposeFx = FxSub(ly, r)
        Case efK:          ComposeFx = FxK(ly, r)
        Case efE:          ComposeFx = FxE(ly, r)
        Case efSep:        ComposeFx = FxSep(ly, r)
        Case efKx:         ComposeFx = FxKx(ly, sp, r)
        Case efBlkx:       ComposeFx = FxBlkx(ly, sp, r)
        Case efCat:        ComposeFx = FxCat(ly, sp, r)
        Case efValue:      ComposeFx = FxValue(ly, sp, band, element, r)
        Case efBase:       ComposeFx = FxBase(ly, sp, band, r, mAggregateOK)
        Case efMask:       ComposeFx = FxMask(ly, sp, r)
        Case efPinValue:   ComposeFx = FxPinValue(ly, sp, band, element, tone, r)
        Case efPinStem:    ComposeFx = FxPinStem(ly, sp, band, element, tone, r)
        Case efPointLabel: ComposeFx = FxPointLabel(ly, sp, st, band, element, tone, r)
    End Select
End Function


' ============================================================================
'  Annotation tables
'
'  Six families of (x, y[, label]) triples. They are what the chart draws its
'  dividers, band rules, panel baselines, titles and tick labels from - all as
'  XY scatter series on the SECONDARY axes, which is what lets them be
'  positioned in category space while the data sits on the primary axes.
'
'  Coordinates are built in (category, value) space so the geometry reads the
'  same for every kind; AnnPut is the only place that knows bar swaps them.
' ============================================================================

Private Sub WriteAnnotations(ByVal ws As Worksheet, ByRef sp As TPanelSpec, _
                             ByRef st As TPanelStyle, ByRef ly As TPanelLayout, _
                             ByRef anns() As TAnnotation, ByRef anCount As Long)

    Dim xs() As Variant, ys() As Variant, labs() As String
    Dim n As Long, cap As Long
    Dim b As Long, band As Long, p As Long, t As Long
    Dim col As Long
    Dim zero As String, gapY As Double

    ReDim anns(1 To AN_MAX)
    anCount = 0
    col = ly.anC0
    cap = 3 * PS_Blocks(sp) * PS_Bands(sp) + 16     ' generous; grown if needed

    ' --- 1. dividers between panel blocks, on the spacer slot, full depth
    AnnStart xs, ys, labs, n, cap
    For b = 0 To PS_Blocks(sp) - 2
        AnnPut sp, xs, ys, labs, n, PS_SpacerCentre(sp, b), CDbl(0), ""
        AnnPut sp, xs, ys, labs, n, PS_SpacerCentre(sp, b), CDbl(PS_Bands(sp)), ""
        AnnPut sp, xs, ys, labs, n, PS_SpacerCentre(sp, b), Empty, ""
    Next b
    If n = 0 Then AnnPut sp, xs, ys, labs, n, 0.5, Empty, ""
    AnnEmit ws, ly, anns, anCount, col, "divider", xs, ys, labs, n, 3, False

    ' --- 2. band rules, in the MIDDLE of the gap between bands, so they never
    '        coincide with a baseline
    AnnStart xs, ys, labs, n, cap
    gapY = PS_FloorPad(sp)
    For band = 1 To PS_Bands(sp) - 1
        AnnPut sp, xs, ys, labs, n, 0.5, band - gapY, ""
        AnnPut sp, xs, ys, labs, n, PS_NSlots(sp) + 0.5, band - gapY, ""
        AnnPut sp, xs, ys, labs, n, PS_NSlots(sp) + 0.5, Empty, ""
    Next band
    If n = 0 Then AnnPut sp, xs, ys, labs, n, 0.5, Empty, ""
    AnnEmit ws, ly, anns, anCount, col, "bandrule", xs, ys, labs, n, 3, False

    ' --- 3. one baseline per panel, at its band's zero line - IF zero is on
    '        the scale at all. A grid whose scale does not include zero (a dot
    '        grid, by default) would otherwise rule a line where zero WOULD be,
    '        which is outside the band and lands in the neighbour's. Written as
    '        a formula rather than decided here, so the rule appears by itself
    '        the day the pasted data crosses zero.
    AnnStart xs, ys, labs, n, cap
    For band = 0 To PS_Bands(sp) - 1
        zero = FxBaselineY(ly, band)
        For b = 0 To PS_Blocks(sp) - 1
            AnnPut sp, xs, ys, labs, n, PS_BlockStart(sp, b) - 0.5, zero, ""
            AnnPut sp, xs, ys, labs, n, PS_BlockEnd(sp, b) + 0.5, zero, ""
            AnnPut sp, xs, ys, labs, n, PS_BlockEnd(sp, b) + 0.5, Empty, ""
        Next b
    Next band
    AnnEmit ws, ly, anns, anCount, col, "baseline", xs, ys, labs, n, 3, False

    ' --- 4. panel titles, each label linked to the input block's name cell
    AnnStart xs, ys, labs, n, cap
    Dim blk As Long, bnd As Long, catV As Variant, valV As Variant
    For p = 0 To PS_NPanels(sp) - 1
        PS_Place sp, p \ sp.cols, p Mod sp.cols, blk, bnd
        If PS_Transposed(sp) Then
            ' Centred in the two-slot gap above the block. The chart layer uses
            ' the Right label position, which centres the text vertically on
            ' its anchor - Above instead lifts it a full text-height and lands
            ' it on the divider.
            catV = PS_BlockEnd(sp, blk) + 1#
            valV = bnd + sp.BandFrac * 0.02
        Else
            catV = CDbl(PS_BlockStart(sp, blk))
            valV = "=" & bnd & "+" & Par(ly, PR_BANDFRAC) & "+0.03"
        End If
        AnnPut sp, xs, ys, labs, n, catV, valV, _
               "='" & SheetRef(ws) & "'!" & _
               "$A$" & (ly.inR0 + p)
    Next p
    AnnEmit ws, ly, anns, anCount, col, "ptitle", xs, ys, labs, n, 1, True

    ' --- 5. tick labels standing in for the hidden value axis, one set per band
    AnnStart xs, ys, labs, n, cap
    Dim f As Double
    For band = 0 To PS_Bands(sp) - 1
        For t = 0 To sp.NTicks - 1
            f = t / CDbl(sp.NTicks - 1)
            If PS_Transposed(sp) Then
                catV = sp.lead * 0.5
            Else
                catV = sp.lead + 0.35
            End If
            AnnPut sp, xs, ys, labs, n, catV, _
                   "=" & band & "+" & NumStr(f) & "*" & Par(ly, PR_BANDFRAC), _
                   FxTickText(ly, f)
        Next t
    Next band
    AnnEmit ws, ly, anns, anCount, col, "ytick", xs, ys, labs, n, 1, True

    ' --- 6. an optional rule inside every panel, after one period. A grid may
    '        change notation part way along - measured up to here, planned
    '        after - and a change of fill alone is not enough at a glance.
    '        Drawn per panel rather than across the chart, because it belongs
    '        to the period axis inside a panel, not to the page.
    If st.dividerAfter > 0 And st.dividerAfter < sp.Periods Then
        AnnStart xs, ys, labs, n, cap
        For band = 0 To PS_Bands(sp) - 1
            For b = 0 To PS_Blocks(sp) - 1
                If PS_PanelAt(sp, b, band) >= 0 Then
                    AnnPut sp, xs, ys, labs, n, _
                           PS_Slot(sp, b, st.dividerAfter, 0) - 0.5, _
                           band + 0.02, ""
                    AnnPut sp, xs, ys, labs, n, _
                           PS_Slot(sp, b, st.dividerAfter, 0) - 0.5, _
                           "=" & band & "+" & Par(ly, PR_BANDFRAC) & "*0.92", ""
                    AnnPut sp, xs, ys, labs, n, _
                           PS_Slot(sp, b, st.dividerAfter, 0) - 0.5, Empty, ""
                End If
            Next b
        Next band
        If n = 0 Then AnnPut sp, xs, ys, labs, n, 0.5, Empty, ""
        AnnEmit ws, ly, anns, anCount, col, "split", xs, ys, labs, n, 3, False
    End If
End Sub


Private Sub AnnStart(ByRef xs() As Variant, ByRef ys() As Variant, _
                     ByRef labs() As String, ByRef n As Long, ByVal cap As Long)
    ReDim xs(1 To cap)
    ReDim ys(1 To cap)
    ReDim labs(1 To cap)
    n = 0
End Sub


' Append one annotation point, in (category, value) space.
'
' THE TRANSPOSED SWAP LIVES HERE AND NOWHERE ELSE. For bar the category axis is
' vertical, so a scatter overlay's x is the value and its y is the category.
'
' `val` is Empty to break a polyline. Note which column the break lands in
' after the swap: for the upright kinds the break is in y and is written as
' NA(); for bar it lands in x and is written as a blank. That asymmetry is
' deliberate and matches the Python - an XValues range must not contain an
' ERROR, and NA() is one.
Private Sub AnnPut(ByRef sp As TPanelSpec, ByRef xs() As Variant, _
                   ByRef ys() As Variant, ByRef labs() As String, _
                   ByRef n As Long, ByVal cat As Variant, ByVal val As Variant, _
                   ByVal label As String)

    n = n + 1
    If n > UBound(xs) Then
        ReDim Preserve xs(1 To n * 2)
        ReDim Preserve ys(1 To n * 2)
        ReDim Preserve labs(1 To n * 2)
    End If

    Dim first As Variant, second As Variant
    If PS_Transposed(sp) Then
        first = val: second = cat
    Else
        first = cat: second = val
    End If

    If IsEmpty(first) Then xs(n) = Empty Else xs(n) = first
    If IsEmpty(second) Then ys(n) = "=NA()" Else ys(n) = second
    labs(n) = label
End Sub


' Write one family and record where it went.
'
' x and y go out as separate assignments, never one rectangle: a mixed array
' assigned to .Formula keeps numbers as numbers and "=..." as formulas, which
' is exactly what is needed, but the two columns carry different kinds of
' content and conflating them is how an NA() ends up in an XValues range.
Private Sub AnnEmit(ByVal ws As Worksheet, ByRef ly As TPanelLayout, _
                    ByRef anns() As TAnnotation, ByRef anCount As Long, _
                    ByRef col As Long, ByVal name As String, _
                    ByRef xs() As Variant, ByRef ys() As Variant, _
                    ByRef labs() As String, ByVal n As Long, _
                    ByVal unit As Long, ByVal hasLabels As Boolean)

    Dim i As Long, a() As Variant
    Dim r1 As Long
    r1 = ly.ptR0 + n - 1

    With ws.Cells(ly.ptR0 - 1, col)
        .value = name
        .Font.bold = True
        .Font.Size = 8
    End With

    ReDim a(1 To n, 1 To 1)
    For i = 1 To n
        a(i, 1) = xs(i)
    Next i
    ws.Range(ws.Cells(ly.ptR0, col), ws.Cells(r1, col)).formula = a

    For i = 1 To n
        a(i, 1) = ys(i)
    Next i
    ws.Range(ws.Cells(ly.ptR0, col + 1), ws.Cells(r1, col + 1)).formula = a

    anCount = anCount + 1
    anns(anCount).name = name
    anns(anCount).col = col
    anns(anCount).r0 = ly.ptR0
    anns(anCount).n = n
    anns(anCount).unit = unit
    anns(anCount).labCol = 0

    Dim width As Long
    width = 2
    If hasLabels Then
        For i = 1 To n
            a(i, 1) = labs(i)
        Next i
        ws.Range(ws.Cells(ly.ptR0, col + 2), ws.Cells(r1, col + 2)).formula = a
        anns(anCount).labCol = col + 2
        width = 3
    End If

    ' One blank column between families, so a stray fill or a widened column
    ' never makes two of them look like one table.
    col = col + width + 1
End Sub


' ============================================================================
'  Helpers
' ============================================================================

' Placeholder values with enough shape to show the chart is working.
'
' Deliberately boring and reproducible: a per-panel level, a mild trend and a
' repeating wobble, with later elements tracking the first. Real data replaces
' it; the point is only that every panel is distinguishable.
Public Function PanelDummyData(ByRef sp As TPanelSpec, _
                               Optional ByVal seed As Long = 7) As Variant
    Dim arr() As Variant
    Dim p As Long, e As Long, k As Long, c As Long
    Dim level As Double, damp As Double, trend As Double, wobble As Double

    ReDim arr(1 To PS_NPanels(sp), 1 To sp.Elements * sp.Periods)

    For p = 0 To PS_NPanels(sp) - 1
        level = 400 + 55 * ((p * 7 + seed) Mod 11)
        c = 0
        For e = 0 To sp.Elements - 1
            If e = 0 Then damp = 1# Else damp = 0.94 + 0.03 * e
            For k = 0 To sp.Periods - 1
                trend = 1 + 0.025 * k
                wobble = IIf(e = 0, 60, 30) * (((k + p) Mod 3) - 1)
                c = c + 1
                arr(p + 1, c) = Round(level * trend * damp + wobble, 1)
            Next k
        Next e
    Next p

    PanelDummyData = arr
End Function


' A sheet name safe to embed in a formula. An apostrophe in the name has to be
' doubled, or a sheet the user renames to "Bob's panel" silently breaks every
' linked label on the chart.
Public Function SheetRef(ByVal ws As Worksheet) As String
    SheetRef = Replace(ws.name, "'", "''")
End Function


' "Panel", or "Panel2", "Panel3"... A build never overwrites an existing sheet,
' which is why this tool needs no ConfirmDestructive prompt.
Public Function UniqueSheetName(ByVal wb As Workbook, _
                                Optional ByVal base As String = "Panel") As String
    Dim nm As String, i As Long
    nm = base
    i = 1
    Do While SheetExists(wb, nm)
        i = i + 1
        nm = base & i
    Loop
    UniqueSheetName = nm
End Function


Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = wb.Sheets(nm)
    SheetExists = (Err.Number = 0)
    Err.Clear
End Function
