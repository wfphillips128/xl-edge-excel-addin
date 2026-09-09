Attribute VB_Name = "modAbout"
Option Explicit

' ============================================================================
'  modAbout
'  Identity of the add-in: its name, its version, and its licence text, plus
'  the ribbon entry point for the About dialog.
'
'  Same split as modSettingsUI: this module holds the FACTS, frmXLEdgeAbout
'  holds the LAYOUT. Two consequences worth knowing --
'
'   * Bumping the version is a one-line edit HERE. Nothing else in the project
'     hardcodes a version number, so there is no second place to forget.
'   * The licence text lives in code, not on shtReference. It is a legal
'     statement that must not be user-editable, which is exactly the opposite
'     of everything AddInStorage manages. Keeping it out of the settings
'     plumbing means no one can blank it out from the settings form.
' ============================================================================

' ----------------------------------------------------------------------------
'  EDIT THIS when you ship a new build.
' ----------------------------------------------------------------------------
Public Const XLEDGE_NAME    As String = "XL Edge"
Public Const XLEDGE_VERSION As String = "2.02"

' ============================================================================
'  Ribbon entry point
' ============================================================================

' Ribbon callback. The `control As IRibbonControl` signature is required by the
' ribbon and matches every other menu macro in this add-in.
Public Sub ShowAboutDialog(control As IRibbonControl)
    LaunchAbout
End Sub

' Callable from the Immediate window or another macro:  modAbout.LaunchAbout
Public Sub LaunchAbout()
    On Error GoTo failed

    ' No FastModeOn here. The dialog is pure UI -- it reads nothing from any
    ' worksheet, so there is nothing to speed up and nothing to protect.
    Dim f As frmXLEdgeAbout
    Set f = New frmXLEdgeAbout          ' explicit instance, not the default one
    f.Show vbModal
    Unload f
    Set f = Nothing
    Exit Sub

failed:
    ' The form is the nice presentation, not the only one. If it fails to build
    ' for any reason, the name and version still reach the user.
    MsgBox VersionLine() & vbCrLf & vbCrLf & _
           "The About window could not be opened." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.description, _
           vbExclamation, "About " & XLEDGE_NAME
End Sub

' ============================================================================
'  Text the dialog displays
' ============================================================================

' "XL Edge   Version 1.12" -- one place that decides how a version is written.
Public Function VersionLine() As String
    VersionLine = XLEDGE_NAME & "   Version " & XLEDGE_VERSION
End Function

Public Function VersionString() As String
    VersionString = "Version " & XLEDGE_VERSION
End Function

' The MIT licence, reproduced verbatim from "MIT License.txt".
'
' Built with one assignment per paragraph rather than a single statement broken
' over many `& _` continuations. VBA caps a statement at 25 continuation lines,
' and a licence is precisely the kind of text that grows past a cap and then
' fails to compile at the worst moment. Appending in steps has no such limit.
'
' Line breaks inside a paragraph are deliberately absent: the dialog's text box
' wraps to its own width, so hard-wrapping here would fight it and produce a
' ragged column on a resized form.
Public Function LicenseText() As String
    Dim s As String

    s = "MIT License" & vbCrLf & vbCrLf

    s = s & "Permission is hereby granted, free of charge, to any person " & _
            "obtaining a copy of this software, to deal in the Software " & _
            "without restriction, including without limitation the rights " & _
            "to use, copy, modify, merge, publish, distribute, sublicense, " & _
            "and/or sell copies of the Software, and to permit persons to " & _
            "whom the Software is furnished to do so, subject to the " & _
            "following conditions:" & vbCrLf & vbCrLf

    s = s & "The above permission notice shall be included in all copies " & _
            "or substantial portions of the Software." & vbCrLf & vbCrLf

    s = s & "THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY " & _
            "KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE " & _
            "WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR " & _
            "PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHOR BE " & _
            "LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN " & _
            "AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT " & _
            "OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER " & _
            "DEALINGS IN THE SOFTWARE."

    LicenseText = s
End Function
