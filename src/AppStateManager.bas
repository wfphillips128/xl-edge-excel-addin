Attribute VB_Name = "AppStateManager"
Option Explicit

' ============================================================================
'  AppStateManager
'  Drop this standard module into the .XLAM add-in ONCE. Then start every
'  menu macro with FastModeOn and end it (on the cleanup label) with
'  FastModeOff.
'
'  Why this and not hand-toggling Application.* in each macro:
'   * It CAPTURES the user's real prior state and RESTORES it, so a user who
'     was already in manual calc isn't yanked back to automatic.
'   * It is RE-ENTRANCY SAFE via a depth counter: if macro A calls macro B,
'     B's FastModeOff won't restore state out from under A. State is only
'     captured on the outermost FastModeOn and only restored on the matching
'     outermost FastModeOff.
'   * Because an add-in changes GLOBAL Excel state (all open workbooks), a
'     crash with EnableEvents = False would otherwise leave events dead for
'     the whole session. Always pair FastModeOff with On Error GoTo CleanExit.
' ============================================================================

Private Type TAppState
    ScreenUpdating  As Boolean
    Calculation     As XlCalculation
    EnableEvents    As Boolean
    DisplayStatusBar As Boolean
    DisplayAlerts   As Boolean
    cursor          As XlMousePointer
End Type

Private mSaved As TAppState
Private mDepth As Long

' Turn on "fast mode". Safe to nest; only the outermost call captures/sets.
Public Sub FastModeOn()
    If mDepth = 0 Then
        With Application
            mSaved.ScreenUpdating = .ScreenUpdating
            mSaved.Calculation = .Calculation
            mSaved.EnableEvents = .EnableEvents
            mSaved.DisplayStatusBar = .DisplayStatusBar
            mSaved.DisplayAlerts = .DisplayAlerts
            mSaved.cursor = .cursor

            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            .EnableEvents = False
            .DisplayStatusBar = False
            .cursor = xlWait
        End With
    End If
    mDepth = mDepth + 1
End Sub

' Restore prior state. Safe to nest; only the outermost call restores.
' Call this from the cleanup label of every macro so it runs even on error.
Public Sub FastModeOff()
    If mDepth = 0 Then Exit Sub          ' nothing to restore
    mDepth = mDepth - 1
    If mDepth > 0 Then Exit Sub          ' still inside an outer FastModeOn

    With Application
        .ScreenUpdating = mSaved.ScreenUpdating
        .Calculation = mSaved.Calculation
        .EnableEvents = mSaved.EnableEvents
        .DisplayStatusBar = mSaved.DisplayStatusBar
        .DisplayAlerts = mSaved.DisplayAlerts
        .cursor = mSaved.cursor
    End With
End Sub

' Emergency reset - bind to a menu button. Use if a hard crash left Excel in
' fast mode (events off, screen frozen) and the depth counter is out of sync.
Public Sub FastModeReset()
    mDepth = 0
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayStatusBar = True
        .DisplayAlerts = True
        .cursor = xlDefault
    End With
End Sub

' Optional: suppress page-break recalculation on a specific sheet during a run.
' Page breaks are a per-WORKSHEET property (not Application), and recalculating
' them is a common hidden cost on big writes. Call before, restore after.
Public Function SuppressPageBreaks(ByVal ws As Worksheet) As Boolean
    SuppressPageBreaks = ws.DisplayPageBreaks   ' return prior value to restore
    ws.DisplayPageBreaks = False
End Function
