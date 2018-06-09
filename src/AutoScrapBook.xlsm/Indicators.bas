Attribute VB_Name = "Indicators"
Sub PutRedFrame()
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
