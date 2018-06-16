Attribute VB_Name = "Indicators"
Public Sub PutCallOut(Optional void = Empty)
    Dim callOutText As String: callOutText = InputBox("吹き出しの内容を入力してください。", "入力")
    If callOutText = "" Then
        MsgBox "キャンセルしました。", vbInformation, "キャンセル"
        GoTo Fin
    End If
    Dim r As Range: Set r _
        = ActiveSheet.Cells(ActiveWindow.ScrollRow, ActiveWindow.ScrollColumn)
    
    Dim sh As Shape
    Set sh = ActiveSheet.Shapes.AddShape( _
        msoShapeRoundedRectangularCallout, r.Left + 10, r.Top + 10, 150, 25)
    sh.Adjustments.Item(1) = -0.1
    sh.Adjustments.Item(2) = 1
    
    Call drawShapeStyle(sh)
    
    sh.TextFrame2.TextRange.Characters.Text = callOutText
    sh.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    sh.Height = sh.Height + 1 '表示バグ解消のためサイズを微修正
Fin:
End Sub

Private Sub drawShapeStyle(ByVal sh As Shape, _
    Optional background_color As Long = rgbLightYellow, _
    Optional font_color As Long = rgbBlack, _
    Optional line_color As Long = rgbBlack)
    
    With sh.Fill
        .Visible = msoTrue
        .ForeColor.RGB = background_color
        .Transparency = 0
        .Solid
    End With
    With sh.line
        .Visible = msoTrue
        .ForeColor.RGB = line_color
        .Transparency = 0
        .Weight = 1.5
    End With
    With sh.TextFrame2.TextRange
        With .Font
            .Name = "MS UI Gothic"
            .Size = 12
            .Bold = msoTrue
            With .Fill
                .Visible = msoTrue
                .ForeColor.RGB = font_color
                .Transparency = 0
                .Solid
            End With
        End With
    End With
End Sub

Public Sub PutRedFrame(Optional void = Empty)
Attribute PutRedFrame.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim r As Range: Set r _
        = ActiveSheet.Cells(ActiveWindow.ScrollRow, ActiveWindow.ScrollColumn)
        
    With ActiveSheet.Shapes.AddShape _
        (msoShapeRectangle, r.Left + 10, r.Top + 10, 50, 50)
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = vbWhite
            .Transparency = 1
            .Solid
        End With
        With .line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(192, 0, 0)
            .Transparency = 0
            .Weight = 2.25
        End With
    End With
End Sub
