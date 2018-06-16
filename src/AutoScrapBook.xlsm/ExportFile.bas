Attribute VB_Name = "ExportFile"
'Require Reference of Microsoft Word Object Library
Public Sub ExportToWord(Optional void = Empty)
    Dim c As New Collection
    
    '文字列のピックアップ
    Dim r As Range
    For Each r In Range(Cells(1, 1), Cells.SpecialCells(xlCellTypeLastCell))
        If r.Value <> "" Then
            With New ParagraphItem
                Set .Item = r
                c.Add .Self
            End With
        End If
    Next
    
    '画像のピックアップ
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
    '出力
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

