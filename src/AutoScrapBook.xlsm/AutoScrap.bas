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


Private running As Boolean
Private targetSheet As Worksheet

Public Sub StartAutoScrap()
    Call Config.LoadConfig
    Set targetSheet = WorksheetSelectionForm.OpenDialog
    If targetSheet Is Nothing Then Exit Sub
    MsgBox "AutoCaptureを開始します。" & vbNewLine & _
        "終了するにはStopボタンをクリックしてください。", vbInformation
    running = True
    Call OnTimeScrap
End Sub

Public Sub StopAutoScrap()
    running = False
    Application.Caption = ""
End Sub

Public Sub OnTimeScrap(Optional ByRef void = Empty)
    If targetSheet Is Nothing Then running = False
    Call animateCaption
    Dim CB As Variant: CB = Application.ClipboardFormats
    Dim TargetRowTop As Single
    If Not running Then GoTo Quit
    If CB(1) <> -1 Then
        Dim i As Long
        For i = 1 To UBound(CB)
            If CB(i) = xlClipboardFormatBitmap Then
                With targetSheet
                    If .Shapes.Count > 0 Then
                        TargetRowTop = lowestShapeEdge
                    Else
                        TargetRowTop = targetSheet.Cells(Config.startRow, 1).Top
                    End If
                    
                    Dim cnt As Long: cnt = 1
                    Do While TargetRowTop > .Cells(cnt, Config.startColumn).Top
                        cnt = cnt + 1
                    Loop
                    
                    If cnt < lastUsedRow Then cnt = lastUsedRow
                    
                    If .Shapes.Count > 0 Or lastUsedRow >= Config.startRow Then cnt = cnt + Config.Margin
                    
                    ActiveWindow.ScrollRow = IIf(cnt = 1, 1, cnt - 1)
                    If Config.InsertTime Then
                        With targetSheet.Cells(cnt, Config.startColumn)
                            .NumberFormatLocal = "hh:mm:ss"
                            .Value = Time
                        End With
                        .Paste Destination:=.Cells(cnt + 2, Config.startColumn)
                    Else
                        .Paste Destination:=.Cells(cnt, Config.startColumn)
                    End If
                    
                    Call popUpWindow
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

Private Sub animateCaption()
    Static state As Boolean
    If running Then
        Application.Caption = IIf(state, "★☆★☆Scrapping☆★☆★", "☆★☆★Scrapping★☆★☆")
        state = Not state
    Else
        Application.Caption = ""
    End If
End Sub

Private Property Get lowestShapeEdge() As Single
    Dim ret As Single: ret = 0
    Dim s As Shape
    For Each s In targetSheet.Shapes
        If ret < (s.Top + s.Height) Then ret = s.Top + s.Height
    Next
    lowestShapeEdge = ret
End Property

Private Property Get lastUsedRow() As Long
    With targetSheet.Cells.SpecialCells(xlCellTypeLastCell)
        Dim j As Long, i As Long
        For j = .Row To 1 Step -1
            For i = .Column To 1 Step -1
                If targetSheet.Cells(j, i).Value <> "" Then
                    lastUsedRow = j
                    Exit Property
                End If
        Next i, j
    End With
End Property

Private Sub popUpWindow()
    Dim fgw As Long: fgw = GetForegroundWindow
    Dim baseWindow As Long: baseWindow = prevWindow
    Call SetWindowPos(Application.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Call SetWindowPos(Application.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    DoEvents
    Sleep 1000
    Call SetWindowPos(Application.hWnd, baseWindow, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Call SetWindowPos(fgw, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Call SetWindowPos(fgw, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Property Get prevWindow() As Long
    Dim ret As Long: ret = Application.hWnd
    
    'Excelは複数のウィンドウで構成されるので、
    '別プロセスになるまで手前へ手前へとretにハンドルを格納しつづける。
    '例えばExcelのひとつ手前にメモ帳が表示されていても、
    '単にGetWindow(Application.hwnd, GW_HWNDPREV)と書くだけではまだExcelの内部ウィンドウがヒットしてしまうので、
    '別プロセスが現れるまでループさせる必要がある。
    Do While GetWindowThreadProcessId(Application.hWnd, 0) = GetWindowThreadProcessId(ret, 0)
        ret = GetWindow(ret, GW_HWNDPREV)
    Loop

    prevWindow = ret
End Property

Private Sub clearClipboard()
    Call OpenClipboard
    Call EmptyClipboard
    Call CloseClipboard
End Sub

