Attribute VB_Name = "modSettingsUI"
Option Explicit

' ============================================================================
'  modSettingsUI
'  Ribbon entry points for the XL Edge settings dialog.
'
'  Deliberately thin: reading and writing storage is AddInStorage's job, laying
'  out and validating the dialog is the form's job, and this module only
'  connects a ribbon button to them. Keeping the launcher out of AddInStorage
'  means the storage module stays free of any UI dependency -- it can be reused
'  or tested without a form in the project.
' ============================================================================

' Ribbon callback. The `control As IRibbonControl` signature is required by the
' ribbon and matches every other menu macro in this add-in.
Public Sub ShowSettingsDialog(control As IRibbonControl)
    LaunchSettings
End Sub

' Callable from the Immediate window or another macro:  modSettingsUI.LaunchSettings
Public Sub LaunchSettings()
    On Error GoTo Failed

    If Not AddInStorage.StorageReady() Then
        MsgBox "XL Edge cannot find its reference sheet, so settings are not " & _
               "available." & vbCrLf & vbCrLf & _
               "The add-in is still running normally on its built-in defaults.", _
               vbExclamation, "XL Edge Settings"
        Exit Sub
    End If

    ' Settings edits write to the add-in workbook, so events must be alive --
    ' do NOT wrap this in AppStateManager.FastModeOn.
    Dim f As frmXLEdgeSettings
    Set f = New frmXLEdgeSettings           ' explicit instance, not the default one
    f.Show vbModal

    Dim didSave As Boolean
    didSave = f.Saved
    Unload f
    Set f = Nothing

    If didSave Then
        Application.StatusBar = "XL Edge settings saved."
        ' OnTime resolves a bare procedure name against the ACTIVE workbook, not
        ' the add-in -- so from an .XLAM it fails to find the macro. Qualifying
        ' it with the add-in's own filename is what makes it resolve.
        Application.OnTime Now + TimeSerial(0, 0, 3), _
            "'" & ThisWorkbook.name & "'!modSettingsUI.ClearStatus"
    End If
    Exit Sub

Failed:
    MsgBox "Could not open XL Edge settings." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.description, vbExclamation, "XL Edge Settings"
End Sub

' Hands the status bar back to Excel after the "saved" message.
Public Sub ClearStatus()
    On Error Resume Next
    Application.StatusBar = False
End Sub

' Menu command: force the next settings read to come off the sheet.
' Only needed if you edited shtReference by hand while a macro had events off.
Public Sub ReloadSettings(control As IRibbonControl)
    AddInStorage.InvalidateCache
    MsgBox "XL Edge settings reloaded from the reference sheet.", _
           vbInformation, "XL Edge Settings"
End Sub
