VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' WithEvents can only be declared in a class module. ThisWorkbook is one, which
' is why this does not need a separate class file.
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

' Fires when Excel loads the add-in. Until App is assigned, none of the event
' handlers below are live -- this one line is what switches them on.
Private Sub Workbook_Open()
    Set App = Application
    modFinancial_Standards.CreateShortcuts
End Sub

' Belt and braces: if the add-in is ticked on in the Add-ins dialog during a
' session that is already running, Workbook_Open may already have gone by.
Private Sub Workbook_AddinInstall()
    Set App = Application
    modFinancial_Standards.CreateShortcuts
End Sub

' Unticked in the Add-ins dialog. Excel keeps running, so the keys have to go
' back or they stay hijacked by an add-in that is no longer loaded.
Private Sub Workbook_AddinUninstall()
    modFinancial_Standards.DeleteShortcuts
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    modFinancial_Standards.DeleteShortcuts
End Sub

Private Sub App_WorkbookActivate(ByVal wb As Workbook)
    RefreshActiveFileList
End Sub

Private Sub App_WorkbookOpen(ByVal wb As Workbook)
    RefreshActiveFileList
End Sub

Private Sub App_NewWorkbook(ByVal wb As Workbook)
    RefreshActiveFileList
End Sub

' Closing the front workbook leaves a different one active, or none at all.
' Either way the list is now wrong.
Private Sub App_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    RefreshActiveFileList
End Sub

' Switching between two windows of the same workbook, or between workbooks by
' clicking their windows, does not always raise WorkbookActivate.
Private Sub App_WindowActivate(ByVal wb As Workbook, ByVal Wn As Window)
    RefreshActiveFileList
End Sub

' One place, wrapped in On Error Resume Next.
'
' An unhandled error inside an Application event handler is unusually nasty: it
' can reset VBA state, and resetting VBA state destroys the IRibbonUI object the
' whole LAMBDA group depends on. A handler that fires on every workbook switch
' has to be incapable of failing.
Private Sub RefreshActiveFileList()
    On Error Resume Next
    modLambdaRibbon.Lambda_ActiveFileChanged
End Sub

