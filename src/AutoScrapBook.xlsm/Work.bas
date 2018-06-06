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
    Dim lo As ILocator
    Set lo = sl
    Debug.Print lo.Top, lo.Bottom, lo.Left, lo.Right
    lo.Locate Range("c14")
End Sub

Sub testRangeLocator()
    Dim rl As RangeLocator: Set rl = New RangeLocator
    rl.SetRange Range("c3:g7")
    Dim lo As ILocator
    Set lo = rl
    Debug.Print lo.Top, lo.Bottom, lo.Left, lo.Right
    lo.Locate Range("d16")
    lo.Locate Range("c3")
End Sub
