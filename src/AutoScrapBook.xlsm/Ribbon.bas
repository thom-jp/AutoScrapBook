Attribute VB_Name = "Ribbon"
Public Sub Ribbon_onLoad(Ribbon As IRibbonUI)
    Ribbon.ActivateTab "AutoCaptureTab"
End Sub

Public Sub RibbonMacros(control As IRibbonControl)
    Application.Run control.Tag, control
End Sub

Private Sub dummy(control As IRibbonControl)
    Application.Run control.ID
End Sub

Private Sub R_StartAutoScrap()
    Call AutoScrap.StartAutoScrap
End Sub

Private Sub R_StopAutoScrap()
    Call AutoScrap.StopAutoScrap
End Sub

Private Sub R_ClearSheet()
    Cleaning.ClearSheet
End Sub

Private Sub R_PutRedFrame()
    Indicators.PutRedFrame
End Sub

Private Sub R_PutCallOut()
    Indicators.PutCallOut
End Sub

Private Sub R_ExportToOtherWorkbook()
    ExportFile.ExportToExcel
End Sub

Private Sub R_AlignScraps()
    Relocation.Main
End Sub

Private Sub R_ExportToWord()
    ExportFile.ExportToWord
End Sub
