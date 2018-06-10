Attribute VB_Name = "ExportFile"
'Require Reference of Microsoft Word Object Library
Sub ExportToWord()
    Dim C As New Collection
    
    '������̃s�b�N�A�b�v
    Dim r As Range
    For Each r In Range(Cells(1, 1), Cells.SpecialCells(xlCellTypeLastCell))
        If r.Value <> "" Then
            With New ParagraphItem
                Set .Item = r
                C.Add .Self
            End With
        End If
    Next
    
    '�摜�̃s�b�N�A�b�v
    Dim s As Shape
    For Each s In ActiveSheet.Shapes
        With New ParagraphItem
            Set .Item = s
            C.Add .Self
        End With
    Next
    
    CSort C, "SortByVerticalLocation"
    
    Dim WD As New Word.Application
    WD.Visible = True
    WD.Documents.Add
    '�o��
    Dim p As ParagraphItem
    For Each p In C
        If IsObject(p.Item) Then
            p.Item.Copy
            WD.Selection.Paste
        Else
            WD.Selection.TypeText p.Item
        End If
        WD.Selection.TypeParagraph
    Next
End Sub

Function SortByVerticalLocation(V As ParagraphItem) As Double
    SortByVerticalLocation = V.Top
End Function

Sub ExportToExcel()
    With ActiveWorkbook
        ActiveSheet.Copy
        .Activate
    End With
End Sub

