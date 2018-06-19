Attribute VB_Name = "ExportFile"
#Const EearlyBinding = False
Private Enum fileType
    xlsx
    docx
End Enum
'Require Reference of Microsoft Word Object Library
Public Sub ExportToWord(Optional void = Empty)
    Dim filePath As String: filePath = GetSavePath(docx)
    If filePath = vbNullString Then Exit Sub
    
    Call Relocation.Main
    DoEvents
    
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
    
    #If EearlyBinding Then
        Dim WD As Word.Application
        Dim doc As Word.Document
        Dim canvas As Word.Shape
    #Else
        Dim WD As Object
        Dim doc As Object
        Dim canvas As Object
        Const wdHorizontalPositionRelativeToPage = 5
        Const wdVerticalPositionRelativeToPage = 6
        Const wdWrapInline = 7
        Const wdAlertsNone = 0
        Const wdAlertsAll = -1
    #End If
    
    Set WD = CreateObject("Word.Application")
    WD.Visible = True
    
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

    '�o��
    Dim p As ParagraphItem
    For Each p In c
        If IsObject(p.Item) Then
            doc.Bookmarks("\EndOfDoc").Select
            x = WD.Selection.Information(wdHorizontalPositionRelativeToPage) + 5
            y = WD.Selection.Information(wdVerticalPositionRelativeToPage) + 5
            Set canvas = doc.Shapes.AddCanvas(x, y, w, h, WD.Selection.Range)
            canvas.WrapFormat.Type = wdWrapInline
            canvas.LockAnchor = True
            With canvas.line
                .Weight = 0.75
                .Style = msoLineSingle
                .Visible = msoTrue
                .ForeColor.RGB = RGB(200, 200, 200)
            End With
            canvas.Fill.BackColor.RGB = RGB(250, 250, 250)
            canvas.Select
            p.Item.CopyPicture
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

    WD.Application.DisplayAlerts = wdAlertsNone
    doc.SaveAs2 filePath
    WD.Application.DisplayAlerts = wdAlertsAll
    
    MsgBox "�o�͂��܂����B", vbInformation, "����"
    Config.LoadConfig
    If Config.Value("CloseAfterExport") Then
        doc.Close
        WD.Application.Quit
    Else
        AppActivate WD.ActiveWindow.Caption & " - " & WD.Caption
    End If
End Sub

#If EearlyBinding Then
Private Sub resizeInsideCanvas(ByRef canvas As Word.Shape)
        Dim n As Word.Shape
#Else
Private Sub resizeInsideCanvas(ByRef canvas As Object)
        Dim n As Object
#End If
    
    canvas.LockAspectRatio = msoFalse
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
    Dim filePath As String: filePath = GetSavePath(xlsx)
    If filePath <> vbNullString Then
        Config.LoadConfig
        ActiveSheet.Copy
        ActiveSheet.Cells.Interior.Color = Config.Value("BackgroundColorForFinish")
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs filePath
        If Config.Value("CloseAfterExport") Then
            ActiveWorkbook.Close
        End If
        Application.DisplayAlerts = True
        MsgBox "�o�͂��܂����B", vbInformation, "����"
    End If
End Sub

Private Function GetSavePath(file_type As fileType) As String
    Dim attr As String
    Dim filter As String
    If file_type = docx Then
        attr = "docx"
        filter = "Word ����, *.docx"
    ElseIf file_type = xlsx Then
        attr = "xlsx"
        filter = "Excel �u�b�N, *.xlsx"
    End If

Retry:
    Dim filePath As Variant
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=Format(Now, "yyyy�Nm��d��_hh��n��s�b") & "_ScreenShots." & attr, _
        FileFilter:=filter)
    
    If filePath <> False Then
        If Dir(filePath) <> "" Then
            Select Case MsgBox("���̃t�@�C���͑��݂��܂��B�㏑�����܂����H", vbYesNoCancel + vbExclamation)
            Case vbCancel
                GetSavePath = vbNullString
                GoTo Fin
            Case vbNo
                GoTo Retry
            Case vbYes
                GoTo ReturnPath
            End Select
        Else
            GoTo ReturnPath
        End If
    Else
        GetSavePath = vbNullString
        GoTo Fin
    End If
ReturnPath:
    GetSavePath = filePath
Fin:
End Function
