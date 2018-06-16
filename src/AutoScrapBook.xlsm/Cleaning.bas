Attribute VB_Name = "Cleaning"
Option Explicit
Public Sub ClearSheet(Optional void = Empty)
    If MsgBox("現在のシートをクリアしますか？", vbYesNo + vbExclamation, "確認") = vbYes Then
        ThisWorkbook.ActiveSheet.Cells.Delete
        Dim sh As Shape
        For Each sh In ThisWorkbook.ActiveSheet.Shapes
            sh.Delete
        Next
    End If
    ActiveWindow.Zoom = 100
    ActiveSheet.Range("A1").Select
End Sub
