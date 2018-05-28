Attribute VB_Name = "Config"
Public BackGroundColor As Long
Public Margin As Long
Public InsertTime As Boolean

Sub LoadConfig(Optional ByRef void = Empty)
    BackGroundColor = ConfigSheet.Range("B2").Value
    Margin = ConfigSheet.Range("B3").Value
    InsertTime = ConfigSheet.Range("B4").Value
    If InsertTime Then Margin = Margin + 1
End Sub
