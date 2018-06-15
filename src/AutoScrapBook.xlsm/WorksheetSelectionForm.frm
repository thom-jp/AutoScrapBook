VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WorksheetSelectionForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "WorksheetSelectionForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "WorksheetSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function OpenDialog() As Worksheet
    Dim ret As Worksheet
    Me.Show
    With ThisWorkbook
        If IsNull(WorksheetList.Value) Then
            Set ret = Nothing
        ElseIf WorksheetList.Selected(0) Then
            Set ret = .ActiveSheet
        ElseIf WorksheetList.Selected(1) Then
            Set ret = .Worksheets.Add(After:=.Worksheets(.Worksheets.Count))
            ret.Cells.Interior.Color = Config.Value("BackGroundColor")
        Else
            Set ret = .Worksheets(WorksheetList.Value)
            ret.Activate
        End If
    End With
    Set OpenDialog = ret
    Unload Me
End Function

Private Sub CancelButton_Click()
    WorksheetList.Value = Null
    Me.Hide
End Sub

Private Sub StartButton_Click()
    If IsNull(WorksheetList.Value) Then
        MsgBox "シートを選択してください。", vbExclamation
    Else
        Me.Hide
    End If
End Sub

Private Sub UserForm_Initialize()
    WorksheetList.AddItem "(現在のシート)"
    WorksheetList.AddItem "(新規作成)"
    Dim ws As Worksheet
    For Each ws In Worksheets
        If UCase(ws.Name) <> "CONFIG" Then
            WorksheetList.AddItem ws.Name
        End If
    Next
    WorksheetList.Value = "(現在のシート)"
    StartButton.SetFocus
End Sub
