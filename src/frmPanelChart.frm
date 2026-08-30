VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPanelChart 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmPanelChart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPanelChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --- source
Private WithEvents optSelection As MSForms.OptionButton
Attribute optSelection.VB_VarHelpID = -1
Private WithEvents optPlaceholder As MSForms.OptionButton
Attribute optPlaceholder.VB_VarHelpID = -1
Private lblSource As MSForms.label

' --- grid
Private WithEvents cboKind As MSForms.ComboBox
Attribute cboKind.VB_VarHelpID = -1
Private WithEvents txtRows As MSForms.TextBox
Attribute txtRows.VB_VarHelpID = -1
Private WithEvents txtCols As MSForms.TextBox
Attribute txtCols.VB_VarHelpID = -1
Private WithEvents txtPeriods As MSForms.TextBox
Attribute txtPeriods.VB_VarHelpID = -1
Private WithEvents txtElements As MSForms.TextBox
Attribute txtElements.VB_VarHelpID = -1
Private lblShape As MSForms.label
Private lblWarn As MSForms.label

' --- labels
Private txtTitle As MSForms.TextBox
Private txtElementNames As MSForms.TextBox
Private txtPeriodLabels As MSForms.TextBox
Private txtSheetName As MSForms.TextBox

' --- appearance
Private chkChartTitle As MSForms.CheckBox
Private chkValueTicks As MSForms.CheckBox
Private chkValueLabels As MSForms.CheckBox
Private chkIncludeZero As MSForms.CheckBox
Private txtBandFrac As MSForms.TextBox
Private txtNTicks As MSForms.TextBox
Private txtLead As MSForms.TextBox
Private txtFloorHead As MSForms.TextBox

Private WithEvents cmdOK As MSForms.CommandButton
Attribute cmdOK.VB_VarHelpID = -1
Private WithEvents cmdCancel As MSForms.CommandButton
Attribute cmdCancel.VB_VarHelpID = -1

' --- state
Private mOK As Boolean
Private mHasBlock As Boolean
Private mNPanels As Long
Private mBuilding As Boolean          ' suppress _Change while BuildLayout runs


' ============================================================================
'  Layout
' ============================================================================

Private Sub UserForm_Initialize()
    mBuilding = True
    BuildLayout
    mBuilding = False
    Revalidate
End Sub


Private Sub BuildLayout()

    Const L As Single = 12            ' left margin
    Const w As Single = 462           ' form width
    Dim y As Single

    Me.caption = "Create Panel Chart"
    Me.width = w
    Me.Height = 520

    ' ---------------------------------------------------------------- source
    y = 8
    Set lblSource = AddLabel("lblSource", L, y, w - 34, 28, "")
    lblSource.WordWrap = True
    y = y + 32

    Set optSelection = Me.Controls.Add("Forms.OptionButton.1", "optSelection", True)
    Place optSelection, L, y, 210, 16
    optSelection.caption = "Use my selection"

    Set optPlaceholder = Me.Controls.Add("Forms.OptionButton.1", "optPlaceholder", True)
    Place optPlaceholder, L + 220, y, 220, 16
    optPlaceholder.caption = "Placeholder data I will paste over"
    y = y + 24

    AddRule L, y, w - 34
    y = y + 8

    ' ------------------------------------------------------------------ grid
    AddLabel "lblKind", L, y + 3, 90, 14, "Chart kind"
    Set cboKind = Me.Controls.Add("Forms.ComboBox.1", "cboKind", True)
    Place cboKind, L + 92, y, 140, 18
    ' A free-text kind is a validation problem invented for no reason.
    cboKind.Style = fmStyleDropDownList
    Dim k As Long
    For k = PK_FIRST To PK_LAST
        cboKind.AddItem KindLabel(k)
    Next k
    cboKind.ListIndex = 0
    y = y + 26

    AddLabel "lblRows", L, y + 3, 90, 14, "Rows"
    Set txtRows = AddText("txtRows", L + 92, y, 44)
    AddLabel "lblCols", L + 150, y + 3, 60, 14, "Columns"
    Set txtCols = AddText("txtCols", L + 212, y, 44)
    y = y + 24

    AddLabel "lblPeriods", L, y + 3, 90, 14, "Periods"
    Set txtPeriods = AddText("txtPeriods", L + 92, y, 44)
    AddLabel "lblElements", L + 150, y + 3, 60, 14, "Elements"
    Set txtElements = AddText("txtElements", L + 212, y, 44)
    y = y + 28

    ' The live readout. It costs nothing - modPanelSpec is pure arithmetic and
    ' touches no COM - and it is the reason a form was chosen over a chain of
    ' InputBoxes: switch the kind from Line to Bar and every limit re-runs
    ' against the TRANSPOSED geometry, with the numbers changing under the
    ' cursor rather than after the build.
    Set lblShape = AddLabel("lblShape", L, y, w - 34, 14, "")
    lblShape.ForeColor = RGB(0, 96, 0)
    y = y + 18

    Set lblWarn = AddLabel("lblWarn", L, y, w - 34, 28, "")
    lblWarn.WordWrap = True
    lblWarn.ForeColor = RGB(160, 80, 0)
    y = y + 32

    AddRule L, y, w - 34
    y = y + 8

    ' ---------------------------------------------------------------- labels
    AddLabel "lblTitle", L, y + 3, 90, 14, "Chart title"
    Set txtTitle = AddText("txtTitle", L + 92, y, w - 130)
    y = y + 24

    AddLabel "lblEN", L, y + 3, 90, 14, "Element names"
    Set txtElementNames = AddText("txtElementNames", L + 92, y, w - 130)
    txtElementNames.ControlTipText = "Comma separated, one per element. " & _
        "Leave blank for Series 1, Series 2..."
    y = y + 24

    AddLabel "lblPL", L, y + 3, 90, 14, "Period labels"
    Set txtPeriodLabels = AddText("txtPeriodLabels", L + 92, y, w - 130)
    txtPeriodLabels.ControlTipText = "Comma separated, one per period. " & _
        "Leave blank to read them from your selection, or for P1, P2..."
    y = y + 24

    AddLabel "lblSheet", L, y + 3, 90, 14, "New sheet name"
    Set txtSheetName = AddText("txtSheetName", L + 92, y, 140)
    y = y + 28

    AddRule L, y, w - 34
    y = y + 8

    ' ------------------------------------------------------------ appearance
    Set chkChartTitle = AddCheck("chkChartTitle", L, y, 130, "Chart title")
    chkChartTitle.value = True
    Set chkValueTicks = AddCheck("chkValueTicks", L + 140, y, 130, "Value tick labels")
    chkValueTicks.value = True
    Set chkValueLabels = AddCheck("chkValueLabels", L + 280, y, 150, "Label every point")
    y = y + 22

    Set chkIncludeZero = AddCheck("chkIncludeZero", L, y, 200, "Scale includes zero")
    ' The only tri-state in the spec, and the one option most likely to be got
    ' wrong with a plain checkbox. Null means "whatever the kind wants" - False
    ' for dot, True for everything else.
    chkIncludeZero.TripleState = True
    chkIncludeZero.value = Null
    chkIncludeZero.ControlTipText = "Third state = let the chart kind decide " & _
        "(a dot grid says 'here', not 'this far from zero', so it does not)"
    y = y + 24

    AddLabel "lblBF", L, y + 3, 90, 14, "Band fill"
    Set txtBandFrac = AddText("txtBandFrac", L + 92, y, 44)
    txtBandFrac.text = "0.82"
    AddLabel "lblNT", L + 150, y + 3, 60, 14, "Ticks"
    Set txtNTicks = AddText("txtNTicks", L + 212, y, 44)
    txtNTicks.text = "3"
    AddLabel "lblLead", L + 272, y + 3, 60, 14, "Lead"
    Set txtLead = AddText("txtLead", L + 334, y, 44)
    txtLead.ControlTipText = "Category slots reserved for the tick labels. " & _
        "Blank or 0 sizes it automatically."
    y = y + 24

    AddLabel "lblFH", L, y + 3, 90, 14, "Floor headroom"
    Set txtFloorHead = AddText("txtFloorHead", L + 92, y, 44)
    txtFloorHead.text = "0"
    txtFloorHead.ControlTipText = "Pad below the scale minimum. Set it to " & _
        "0.10 when labelling negative points, or their labels print through " & _
        "the rule beneath the band."
    y = y + 32

    ' ---------------------------------------------------------------- buttons
    Set cmdOK = Me.Controls.Add("Forms.CommandButton.1", "cmdOK", True)
    Place cmdOK, w - 200, y, 80, 22
    ' Captioned Create, and deliberately NOT .Default: with a default button,
    ' Enter after typing a value builds the chart when the user meant "apply
    ' this edit".
    cmdOK.caption = "Create"

    Set cmdCancel = Me.Controls.Add("Forms.CommandButton.1", "cmdCancel", True)
    Place cmdCancel, w - 110, y, 80, 22
    cmdCancel.caption = "Cancel"
    cmdCancel.Cancel = True

    Me.Height = y + 68
End Sub


' ============================================================================
'  Preload - everything the launcher knows before the user sees the form
'
'  All scalars. See the header note on why nothing here is a Type.
' ============================================================================

Public Sub Preload(ByVal srcNote As String, ByVal hasBlock As Boolean, _
                   ByVal p As Long, ByVal e As Long, ByVal nPanels As Long, _
                   ByVal confident As Boolean, ByVal sheetGuess As String)

    mBuilding = True
    mHasBlock = hasBlock
    mNPanels = nPanels

    lblSource.caption = srcNote
    txtSheetName.text = sheetGuess

    optSelection.Enabled = hasBlock
    optSelection.value = hasBlock
    optPlaceholder.value = Not hasBlock

    If hasBlock Then
        txtPeriods.text = CStr(p)
        txtElements.text = CStr(e)
        ' Rows and Cols are left BLANK on purpose. The panel count is fixed by
        ' the selection, but 20 panels is equally 5x4, 4x5, 10x2 or 2x10 and
        ' only the user knows which one reads right. Nothing is guessed.
        txtRows.text = ""
        txtCols.text = ""
        If Not confident Then
            lblSource.caption = lblSource.caption & _
                "  The number of periods was not obvious from your headers - " & _
                "check it before continuing."
        End If
    Else
        txtRows.text = "5"
        txtCols.text = "4"
        txtPeriods.text = "12"
        txtElements.text = "1"
    End If

    mBuilding = False
    Revalidate
End Sub


' ============================================================================
'  Validation - one function, run from every keystroke and again on Create
'
'  It builds a real TPanelSpec and hands it to modPanelSpec.PanelSpecValidate,
'  which is the SAME validator the Immediate-window entry point uses. There is
'  no second place for the rules to drift from.
' ============================================================================

Private Sub Revalidate()

    If mBuilding Then Exit Sub

    Dim sp As TPanelSpec
    Dim problem As String, warn As String

    On Error GoTo Broken

    sp.kind = KindIndex
    sp.rows = AsLong(txtRows.text)
    sp.cols = AsLong(txtCols.text)
    sp.Periods = AsLong(txtPeriods.text)
    sp.Elements = AsLong(txtElements.text)
    sp.BandFrac = AsDouble(txtBandFrac.text, 0.82)
    sp.NTicks = AsLong(txtNTicks.text)
    If sp.NTicks = 0 Then sp.NTicks = 3
    sp.lead = AsLong(txtLead.text)
    sp.FloorHeadroom = AsDouble(txtFloorHead.text, 0)

    If sp.rows < 1 Or sp.cols < 1 Then
        lblShape.caption = ""
        If mHasBlock Then
            lblWarn.caption = mNPanels & " panel" & IIf(mNPanels = 1, "", "s") & _
                " selected. Enter a grid that holds at least " & mNPanels & "."
        Else
            lblWarn.caption = "Enter the grid size."
        End If
        cmdOK.Enabled = False
        Exit Sub
    End If

    problem = PanelSpecValidate(sp)
    If Len(problem) = 0 Then
        lblShape.caption = PanelSpecDescribe(sp)
    Else
        lblShape.caption = ""
    End If

    ' Refusing a grid that would silently DROP panels. Allowing one that holds
    ' more cells than panels: the engine draws nothing and prints no title
    ' where a panel is absent, and for a prime panel count there is no factor
    ' pair at all.
    If Len(problem) = 0 And mHasBlock And optSelection.value Then
        If sp.rows * sp.cols < mNPanels Then
            problem = "That grid holds " & (sp.rows * sp.cols) & " panels but " & _
                      mNPanels & " are selected - " & _
                      (mNPanels - sp.rows * sp.cols) & " would be dropped."
        End If
    End If

    If Len(problem) = 0 Then problem = SheetNameProblem(txtSheetName.text)

    If Len(problem) > 0 Then
        lblWarn.ForeColor = RGB(192, 0, 0)
        lblWarn.caption = problem
        cmdOK.Enabled = False
        Exit Sub
    End If

    warn = PanelSpecWarnings(sp)
    If mHasBlock And optSelection.value Then
        If sp.rows * sp.cols > mNPanels Then
            If Len(warn) > 0 Then warn = warn & vbCrLf
            warn = warn & (sp.rows * sp.cols) & " cells, " & _
                   (sp.rows * sp.cols - mNPanels) & " will be blank."
        End If
    End If

    lblWarn.ForeColor = RGB(160, 80, 0)
    lblWarn.caption = warn
    cmdOK.Enabled = True
    Exit Sub

Broken:
    lblShape.caption = ""
    lblWarn.ForeColor = RGB(192, 0, 0)
    lblWarn.caption = "Check the numbers."
    cmdOK.Enabled = False
End Sub


Private Function SheetNameProblem(ByVal nm As String) As String
    Dim bad As Variant, i As Long
    nm = Trim$(nm)
    If Len(nm) = 0 Then SheetNameProblem = "The new sheet needs a name.": Exit Function
    If Len(nm) > 31 Then SheetNameProblem = "A sheet name is at most 31 characters.": Exit Function
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(bad) To UBound(bad)
        If InStr(1, nm, CStr(bad(i)), vbTextCompare) > 0 Then
            SheetNameProblem = "A sheet name cannot contain  :  \  /  ?  *  [  ]"
            Exit Function
        End If
    Next i
End Function


' ============================================================================
'  Events
' ============================================================================

Private Sub cboKind_Change()
    Revalidate
End Sub

Private Sub txtRows_Change()
    Revalidate
End Sub

Private Sub txtCols_Change()
    Revalidate
End Sub

Private Sub txtPeriods_Change()
    Revalidate
End Sub

Private Sub txtElements_Change()
    Revalidate
End Sub

Private Sub optSelection_Change()
    Revalidate
End Sub

Private Sub optPlaceholder_Change()
    Revalidate
End Sub

Private Sub cmdOK_Click()
    Revalidate
    If Not cmdOK.Enabled Then Exit Sub
    mOK = True
    ' Hide, never Unload. Unloading destroys these variables and the launcher
    ' would have nothing left to read.
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    mOK = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' The X button is a cancel, not a crash.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        mOK = False
        Me.Hide
    End If
End Sub


' ============================================================================
'  Results - one scalar per field
' ============================================================================

Public Property Get OK() As Boolean
    OK = mOK
End Property

Public Property Get useSelection() As Boolean
    useSelection = optSelection.value
End Property

Public Property Get KindIndex() As Long
    KindIndex = cboKind.ListIndex + PK_FIRST
End Property

Public Property Get GridRows() As Long
    GridRows = AsLong(txtRows.text)
End Property

Public Property Get GridCols() As Long
    GridCols = AsLong(txtCols.text)
End Property

Public Property Get Periods() As Long
    Periods = AsLong(txtPeriods.text)
End Property

Public Property Get Elements() As Long
    Elements = AsLong(txtElements.text)
End Property

Public Property Get SheetName() As String
    SheetName = Trim$(txtSheetName.text)
End Property

Public Property Get TitleText() As String
    TitleText = Trim$(txtTitle.text)
End Property

Public Property Get ElementNamesCsv() As String
    ElementNamesCsv = Trim$(txtElementNames.text)
End Property

Public Property Get PeriodLabelsCsv() As String
    PeriodLabelsCsv = Trim$(txtPeriodLabels.text)
End Property

Public Property Get WantChartTitle() As Boolean
    WantChartTitle = CBool(chkChartTitle.value)
End Property

Public Property Get WantValueTicks() As Boolean
    WantValueTicks = CBool(chkValueTicks.value)
End Property

Public Property Get WantValueLabels() As Boolean
    WantValueLabels = CBool(chkValueLabels.value)
End Property

Public Property Get BandFrac() As Double
    BandFrac = AsDouble(txtBandFrac.text, 0.82)
End Property

Public Property Get NTicks() As Long
    NTicks = AsLong(txtNTicks.text)
End Property

Public Property Get LeadSlots() As Long
    LeadSlots = AsLong(txtLead.text)
End Property

Public Property Get FloorHeadroom() As Double
    FloorHeadroom = AsDouble(txtFloorHead.text, 0)
End Property

' -1 = whatever the kind wants, 0 = no, 1 = yes.
Public Property Get IncludeZeroMode() As Long
    If IsNull(chkIncludeZero.value) Then
        IncludeZeroMode = -1
    ElseIf chkIncludeZero.value Then
        IncludeZeroMode = 1
    Else
        IncludeZeroMode = 0
    End If
End Property


' ============================================================================
'  Control factory
' ============================================================================

Private Sub Place(ByVal c As Object, ByVal x As Single, ByVal y As Single, _
                  ByVal w As Single, ByVal h As Single)
    c.Left = x: c.Top = y: c.width = w: c.Height = h
End Sub

Private Function AddLabel(ByVal nm As String, ByVal x As Single, ByVal y As Single, _
                          ByVal w As Single, ByVal h As Single, _
                          ByVal caption As String) As MSForms.label
    Dim c As MSForms.label
    Set c = Me.Controls.Add("Forms.Label.1", nm, True)
    Place c, x, y, w, h
    c.caption = caption
    Set AddLabel = c
End Function

Private Function AddText(ByVal nm As String, ByVal x As Single, ByVal y As Single, _
                         ByVal w As Single) As MSForms.TextBox
    Dim c As MSForms.TextBox
    Set c = Me.Controls.Add("Forms.TextBox.1", nm, True)
    Place c, x, y, w, 18
    Set AddText = c
End Function

Private Function AddCheck(ByVal nm As String, ByVal x As Single, ByVal y As Single, _
                          ByVal w As Single, ByVal caption As String) As MSForms.CheckBox
    Dim c As MSForms.CheckBox
    Set c = Me.Controls.Add("Forms.CheckBox.1", nm, True)
    Place c, x, y, w, 16
    c.caption = caption
    Set AddCheck = c
End Function

' A one-pixel label standing in for a horizontal rule. MSForms has no
' separator control, and a Frame around every group would cost more height
' than the whole form has.
Private Sub AddRule(ByVal x As Single, ByVal y As Single, ByVal w As Single)
    Static n As Long
    n = n + 1
    Dim c As MSForms.label
    Set c = Me.Controls.Add("Forms.Label.1", "lblRule" & n, True)
    Place c, x, y, w, 1
    c.BackColor = RGB(200, 200, 200)
    c.BackStyle = fmBackStyleOpaque
    c.caption = ""
End Sub


' ============================================================================
'  Parsing
'
'  A blank box is zero, not an error: Rows and Cols START blank when building
'  from a selection, and Revalidate has to be able to say "enter the grid size"
'  rather than "check the numbers".
' ============================================================================

Private Function AsLong(ByVal s As String) As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    AsLong = CLng(Int(val(s)))
End Function

Private Function AsDouble(ByVal s As String, ByVal dflt As Double) As Double
    s = Trim$(s)
    If Len(s) = 0 Then AsDouble = dflt: Exit Function
    If Not IsNumeric(s) Then AsDouble = dflt: Exit Function
    AsDouble = CDbl(s)
End Function



