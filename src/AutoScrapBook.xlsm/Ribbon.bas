Attribute VB_Name = "Ribbon"
Sub Ribbon_onLoad(Ribbon As IRibbonUI)
    Ribbon.ActivateTab "AutoCaptureTab"
    Debug.Print "Loaded"
End Sub

Sub RibbonMacros(control As IRibbonControl)
    Application.Run control.Tag, control
End Sub

Private Sub dummy(control As IRibbonControl)
    Application.Run control.ID
End Sub

Sub StartCapture()
    MsgBox "StartCapture"
End Sub

Sub StopCapture()
    MsgBox "StopCapture"
End Sub

Sub ClearAll()
    MsgBox "Clear"
End Sub

Sub InsertScope()
    MsgBox "PutScope"
End Sub

Sub ExportManual()
    MsgBox "ExportManual"
End Sub

Sub AlignScreenShots()
    MsgBox "AlignScreenShots"
End Sub
