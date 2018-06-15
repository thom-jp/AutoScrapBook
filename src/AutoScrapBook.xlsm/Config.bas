Attribute VB_Name = "Config"
Option Explicit
Public BackGroundColor As Long
Public Margin As Long
Public InsertTime As Boolean
Public startRow As Long
Public startColumn As Long
Private configurations As Variant

Sub LoadConfig(Optional ByRef void = Empty)
    configurations = ThisWorkbook.Worksheets("Config").Range("a1").CurrentRegion.Value
End Sub

Public Property Get Value(conf_name As String)
    Dim i As Long
    For i = LBound(configurations, 1) To UBound(configurations, 1)
        If UCase(conf_name) = UCase(configurations(i, 1)) Then
            Value = configurations(i, 2)
        End If
    Next
End Property
