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
    With ActiveSheet
        Dim sh As Worksheet: Set sh = Worksheets.Add
        .Activate
    End With
    Dim rl As RangeLocator: Set rl = New RangeLocator
    rl.SetRange Sheet1.Range("c3:g7")
    Dim lo As ILocator
    Set lo = rl
    Debug.Print lo.Top, lo.Bottom, lo.Left, lo.Right
    lo.Locate sh.Range("d16")
    lo.Locate Sheet1.Range("c3")
    
    With New NoAlert
        sh.Delete
    End With
End Sub
