Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
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

