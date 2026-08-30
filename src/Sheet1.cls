VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal target As Range)
    ' Only bother when the change actually landed inside the settings table.
    On Error Resume Next
    If Not Intersect(target, Me.ListObjects("tblConstants").Range) Is Nothing Then
        AddInStorage.InvalidateCache
    End If
    If Not Intersect(target, Me.ListObjects("tblCompany").Range) Is Nothing Then
        AddInStorage.InvalidateCache
    End If
End Sub

