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
    Dim doc As Word.Document
    Set doc = WD.Documents.Add
    
    With doc.PageSetup
        .LeftMargin = 54
        .RightMargin = 54
        .BottomMargin = 72
        .TopMargin = 72
    End With
    
    Dim x As Double, y As Double, w As Double, h As Double
    With WD.Selection.PageSetup
        w = .PageWidth - .LeftMargin - .RightMargin
        h = .PageHeight - .TopMargin - .BottomMargin
    End With

    '出力
    Dim p As ParagraphItem
    For Each p In c
        If IsObject(p.Item) Then
            doc.Bookmarks("\EndOfDoc").Select
            x = WD.Selection.Information(Word.wdHorizontalPositionRelativeToPage) + 5
            y = WD.Selection.Information(Word.wdVerticalPositionRelativeToPage) + 5
            Dim canvas As Word.Shape
            Set canvas = doc.Shapes.AddCanvas(x, y, w, h, WD.Selection.Range)
            canvas.WrapFormat.Type = Word.wdWrapInline
            canvas.LockAnchor = True
            With canvas.line
                .Weight = 0.75
                .Style = msoLineSingle
                .Visible = msoTrue
                .ForeColor.RGB = RGB(200, 200, 200)
            End With
            canvas.Fill.BackColor.RGB = RGB(250, 250, 250)
            canvas.Select
            p.Item.Copy
            WD.Selection.Paste
            Call resizeInsideCanvas(canvas)
            doc.Bookmarks("\EndOfDoc").Select
            WD.Selection.TypeParagraph
        Else
            doc.Bookmarks("\EndOfDoc").Select
            WD.Selection.TypeText p.Item
        End If
        doc.Bookmarks("\EndOfDoc").Select
        WD.Selection.TypeParagraph
    Next
End Sub

Private Sub resizeInsideCanvas(ByRef canvas As Word.Shape)
    canvas.LockAspectRatio = msoFalse
    Dim n As Word.Shape
    Set n = canvas.CanvasItems(1)
    If n.Height = canvas.Height And n.Width < canvas.Width Then
        canvas.Height = canvas.Width / (n.Width / n.Height)
        n.Width = canvas.Width
        n.Height = canvas.Height
        n.LockAspectRatio = msoTrue
        n.Width = n.Width * 0.95
        n.Top = 0.2
        n.Left = 0.2
    ElseIf n.Height < canvas.Height Then
        canvas.Height = n.Height + 5
        n.LockAspectRatio = msoFalse
        n.Height = canvas.Height - 5
    End If
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

