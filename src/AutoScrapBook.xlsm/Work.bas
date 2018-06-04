Attribute VB_Name = "Work"
Option Explicit
Sub ClearSheet()
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

Sub testShapeLocator()
    Dim sl As ShapeLocator: Set sl = New ShapeLocator
    sl.SetShape Sheet1.Shapes(1)
    Debug.Print sl.Top, sl.Bottom, sl.Left, sl.Right
End Sub
