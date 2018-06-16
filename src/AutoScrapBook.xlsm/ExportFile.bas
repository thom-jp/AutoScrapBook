Attribute VB_Name = "ExportFile"
'Require Reference of Microsoft Word Object Library
Public Sub ExportToWord(Optional void = Empty)
    Dim c As New Collection
    
    '������̃s�b�N�A�b�v
    Dim r As Range
    For Each r In Range(Cells(1, 1), Cells.SpecialCells(xlCellTypeLastCell))
        If r.Value <> "" Then
            With New ParagraphItem
                Set .Item = r
                c.Add .Self
            End With
        End If
    Next
    
    '�摜�̃s�b�N�A�b�v
    Dim s As Shape
    For Each s In ActiveSheet.Shapes
        With New ParagraphItem
            Set .Item = s
            c.Add .Self
        End With
    Next
    
    CSort c, "SortByVerticalLocation"
    
    Dim WD As New Word.Application
    WD.Visible = True
    WD.Documents.Add
    '�o��
    Dim p As ParagraphItem
    For Each p In c
        If IsObject(p.Item) Then
            p.Item.Copy
            WD.Selection.Paste
        Else
            WD.Selection.TypeText p.Item
        End If
        WD.Selection.TypeParagraph
    Next
End Sub

Public Function SortByVerticalLocation(V As ParagraphItem) As Double
    SortByVerticalLocation = V.Top
End Function

Public Sub ExportToExcel(Optional void = Empty)
    With ActiveWorkbook
        ActiveSheet.Copy
        .Activate
    End With
End Sub

