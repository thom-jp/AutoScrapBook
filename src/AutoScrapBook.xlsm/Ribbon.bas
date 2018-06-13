Attribute VB_Name = "Ribbon"
Sub Ribbon_onLoad(Ribbon As IRibbonUI)
    Ribbon.ActivateTab "AutoCaptureTab"
End Sub

Sub RibbonMacros(control As IRibbonControl)
    Application.Run control.Tag, control
End Sub

Private Sub dummy(control As IRibbonControl)
    Application.Run control.ID
End Sub

Sub R_StartAutoScrap()
    Call AutoScrap.StartAutoScrap
End Sub

Sub R_StopAutoScrap()
    Call AutoScrap.StopAutoScrap
End Sub

Sub R_ClearAll()
    Cleaning.ClearSheet
End Sub

Sub R_PutRedFrame()
    Indicators.PutRedFrame
End Sub

Sub R_PutCallOut()
    Indicators.PutCallOut
End Sub

Sub R_ExportToOtherWorkbook()
    ExportFile.ExportToExcel
End Sub

Sub R_AlignScraps()
    Relocation.Main
End Sub

Sub R_ExportToWord()
    ExportFile.ExportToWord
End Sub
