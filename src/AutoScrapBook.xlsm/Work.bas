Attribute VB_Name = "Work"
Sub ClearSheet()
    If MsgBox("���݂̃V�[�g���N���A���܂����H", vbYesNo + vbExclamation, "�m�F") = vbYes Then
        ThisWorkbook.ActiveSheet.Cells.Delete
        For Each sh In ThisWorkbook.ActiveSheet.Shapes
            sh.Delete
        Next
    End If
    ActiveWindow.Zoom = 100
    ActiveSheet.Range("A1").Select
End Sub
