Attribute VB_Name = "Config"
Option Explicit
Public BackGroundColor As Long
Public Margin As Long
Public InsertTime As Boolean
Public StartRow As Long
Public StartColumn As Long

Sub LoadConfig(Optional ByRef void = Empty)
    BackGroundColor = ConfigSheet.Range("B2").Value
    Margin = ConfigSheet.Range("B3").Value
    InsertTime = ConfigSheet.Range("B4").Value
    If InsertTime Then Margin = Margin + 1
    StartColumn = ConfigSheet.Range("B6").Value
    StartRow = ConfigSheet.Range("B5").Value
End Sub
