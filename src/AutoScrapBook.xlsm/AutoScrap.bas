Attribute VB_Name = "AutoScrap"
Option Explicit
Private Declare Function OpenClipboard Lib "user32" (Optional ByVal hWnd As Long = 0) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Running As Boolean
Private targetSheet As Worksheet

Private Sub clearClipboard()
    Call OpenClipboard
    Call EmptyClipboard
    Call CloseClipboard
End Sub

Private Sub animateCaption()
    Static state As Boolean
    If Running Then
        Application.Caption = IIf(state, "★☆★☆Scrapping☆★☆★", "☆★☆★Scrapping★☆★☆")
        state = Not state
    Else
        Application.Caption = ""
    End If
End Sub

Public Sub StartAutoScrap()
    Call Config.LoadConfig
    Set targetSheet = WorksheetSelectionForm.OpenDialog
    If targetSheet Is Nothing Then Exit Sub
    MsgBox "AutoCaptureを開始します。" & vbNewLine & _
        "終了するにはStopボタンをクリックしてください。", vbInformation
    Running = True
    Call OnTimeScrap
End Sub

Public Sub StopAutoScrap()
    Running = False
    Application.Caption = ""
End Sub

Public Sub OnTimeScrap(Optional ByRef void = Empty)
    If targetSheet Is Nothing Then Running = False
    Call animateCaption
    Dim CB As Variant: CB = Application.ClipboardFormats
    Dim TargetRowTop As Single
    If Not Running Then GoTo Quit
    If CB(1) <> -1 Then
        Dim i As Long
        For i = 1 To UBound(CB)
            If CB(i) = xlClipboardFormatBitmap Then
                With targetSheet
                    If .Shapes.Count > 0 Then
                        With .Shapes(.Shapes.Count)
                            TargetRowTop = .Top + .Height
                        End With
                    Else
                        TargetRowTop = 0
                    End If
                    Dim cnt As Long
                    cnt = 1
                    Do While TargetRowTop > .Cells(cnt, 1).Top
                        cnt = cnt + 1
                    Loop
                    cnt = cnt + Config.Margin
                    ActiveWindow.ScrollRow = cnt - 1
                    If Config.InsertTime Then targetSheet.Range("A" & cnt - 1).Value = Time
                    .Paste Destination:=.Cells(cnt, 1)
                End With
                Call clearClipboard
            End If
        Next i
    End If
    DoEvents
    Application.OnTime DateAdd("s", 1, Now), "OnTimeScrap"
    Exit Sub
Quit:
    MsgBox "AutoScrapを停止しました。", vbInformation
    Application.Caption = ""
End Sub
