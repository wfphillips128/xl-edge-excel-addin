VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmXLEdgeAbout 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmXLEdgeAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmXLEdgeAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
'  Controls, created at run time in UserForm_Initialize.
'
'  WithEvents is what lets a control created in code still fire Click events.
'  A control dragged on in the designer gets that wiring for free; one added
'  with Controls.Add does not, unless it is assigned to a WithEvents variable.
'  Only the button needs it -- labels and the licence box are never clicked.
' ----------------------------------------------------------------------------
Private lblTitle     As MSForms.label
Private lblVersion   As MSForms.label
Private lblLicHdr    As MSForms.label
Private txtLicense   As MSForms.TextBox

Private WithEvents cmdClose As MSForms.CommandButton
Attribute cmdClose.VB_VarHelpID = -1

' Layout constants, so the panel and the button stay aligned if you move one.
Private Const MARGIN_L   As Single = 14
Private Const BODY_W     As Single = 408
Private Const BTN_W      As Single = 84
Private Const BTN_H      As Single = 24

' ============================================================================
'  Setup
' ============================================================================

Private Sub UserForm_Initialize()
    BuildLayout
    FillText
End Sub

Private Sub BuildLayout()
    Me.caption = "About " & XLEDGE_NAME
    Me.width = 436
    Me.Height = 386

    Dim c As Object

    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle", True)
    With lblTitle
        .Left = MARGIN_L: .Top = 12: .width = BODY_W: .Height = 22
        .Font.Size = 14
        .Font.bold = True
    End With

    Set lblVersion = Me.Controls.Add("Forms.Label.1", "lblVersion", True)
    With lblVersion
        .Left = MARGIN_L: .Top = 36: .width = BODY_W: .Height = 14
    End With

    ' A 1-point-tall label with a fill is the cheapest horizontal rule MSForms
    ' offers -- there is no line control on a UserForm.
    Set c = Me.Controls.Add("Forms.Label.1", "lblRule", True)
    With c
        .caption = ""
        .Left = MARGIN_L: .Top = 58: .width = BODY_W: .Height = 1
        .BackColor = &H80000010                ' system button-shadow grey
    End With

    Set lblLicHdr = Me.Controls.Add("Forms.Label.1", "lblLicHdr", True)
    With lblLicHdr
        .caption = "License"
        .Left = MARGIN_L: .Top = 68: .width = BODY_W: .Height = 14
        .Font.bold = True
    End With

    Set txtLicense = Me.Controls.Add("Forms.TextBox.1", "txtLicense", True)
    With txtLicense
        .Left = MARGIN_L: .Top = 86: .width = BODY_W: .Height = 224
        .Multiline = True
        .WordWrap = True
        .ScrollBars = fmScrollBarsVertical
        ' Locked, not Enabled = False. A locked box still lets the user select
        ' and Ctrl-C the licence text -- which is the point of showing it --
        ' while a disabled one greys it out and blocks copying.
        .Locked = True
        .TabStop = False
    End With

    Set cmdClose = Me.Controls.Add("Forms.CommandButton.1", "cmdClose", True)
    With cmdClose
        .caption = "Close"
        .Left = MARGIN_L + BODY_W - BTN_W: .Top = 322: .width = BTN_W: .Height = BTN_H
        .Default = True                        ' Enter closes
        .Cancel = True                         ' Esc closes
    End With
End Sub

' Every string on the form comes from modAbout -- see the header note.
Private Sub FillText()
    lblTitle.caption = XLEDGE_NAME
    lblVersion.caption = modAbout.VersionString()
    txtLicense.text = modAbout.LicenseText()

    ' Put the caret at the top. Assigning .Text leaves the selection at the end
    ' of the string, which on a scrolling box means it opens showing the LAST
    ' paragraph -- the disclaimer -- instead of the first.
    '
    ' On Error Resume Next because SelStart touches a control that has not been
    ' painted yet; cosmetic scroll position is never worth failing the form for.
    On Error Resume Next
    txtLicense.SelStart = 0
    On Error GoTo 0
End Sub

' ============================================================================
'  Events
' ============================================================================

Private Sub cmdClose_Click()
    Unload Me
End Sub

