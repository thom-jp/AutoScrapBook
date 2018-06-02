Attribute VB_Name = "AutoScrap"
Option Explicit
Private Declare Function OpenClipboard Lib "user32" (Optional ByVal hWnd As Long = 0) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal uFlags As Long _
    ) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_HWNDPREV = 3

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long


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
                        TargetRowTop = targetSheet.Cells(Config.StartRow, 1).Top
                    End If
                    Dim cnt As Long
                    cnt = 1
                    Do While TargetRowTop > .Cells(cnt, Config.StartColumn).Top
                        cnt = cnt + 1
                    Loop
                    cnt = cnt + Config.Margin
                    ActiveWindow.ScrollRow = cnt - 1
                    If Config.InsertTime Then
                        With targetSheet.Cells(cnt - 1, Config.StartColumn)
                            .NumberFormatLocal = "hh:mm:ss"
                            .Value = Time
                        End With
                    End If
                    .Paste Destination:=.Cells(cnt, Config.StartColumn)
                    Dim fgw As Long: fgw = GetForegroundWindow
                    Call PopUpWindow
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

Private Sub PopUpWindow()
    Dim fgw As Long: fgw = GetForegroundWindow
    Dim baseWindow As Long: baseWindow = PrevWindow
    Call SetWindowPos(Application.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Call SetWindowPos(Application.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    DoEvents
    Sleep 1000
    Call SetWindowPos(Application.hWnd, baseWindow, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Call SetWindowPos(fgw, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Call SetWindowPos(fgw, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub


Private Property Get PrevWindow() As Long
    Dim ret As Long: ret = Application.hWnd
    
    'Excelは複数のウィンドウで構成されるので、
    '別プロセスになるまで手前へ手前へとretにハンドルを格納しつづける。
    '例えばExcelのひとつ手前にメモ帳が表示されていても、
    '単にGetWindow(Application.hwnd, GW_HWNDPREV)と書くだけではまだExcelの内部ウィンドウがヒットしてしまうので、
    '別プロセスが現れるまでループさせる必要がある。
    Do While GetWindowThreadProcessId(Application.hWnd, 0) = GetWindowThreadProcessId(ret, 0)
        ret = GetWindow(ret, GW_HWNDPREV)
    Loop

    PrevWindow = ret
End Property
