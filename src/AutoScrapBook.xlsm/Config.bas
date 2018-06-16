Attribute VB_Name = "Config"
Option Explicit
Private configurations As Variant   '<- configulations is keepd here as variant array.

Private Property Get InitialSettings() As Collection
    Set InitialSettings = New Collection
    InitialSettings.Add Array("Name", "Value", "Description")   '<- Header.
    
    'Please define settings as like examples below.
    'The examples should be deleted.
    InitialSettings.Add Array("BackgroundColor", rgbWheat, "rgbWheat")
    InitialSettings.Add Array("Margin", 5, "how many blank cells put between pictures")
    InitialSettings.Add Array("InsertTime", True, "Write scraptime or not")
    InitialSettings.Add Array("StartRow", 5, "")
    InitialSettings.Add Array("StartColumn", 3, "")
End Property

Public Sub LoadConfig(Optional ByRef void = Empty)
    If Not existConfigSheet Then ResetConfig
    configurations = ThisWorkbook.Worksheets("Config").Range("a1").CurrentRegion.Value
End Sub

Public Sub ShowConfig(Optional void = Empty)
    If Not existConfigSheet Then ResetConfig
    With ThisWorkbook.Worksheets("Config")
        .Visible = xlSheetVisible
        .Activate
    End With
End Sub

Public Sub HideConfig(Optional void = Empty)
    If Not existConfigSheet Then ResetConfig
    ThisWorkbook.Worksheets("Config").Visible = xlSheetVeryHidden
End Sub

Public Property Get Value(conf_name As String)
    Dim i As Long
    For i = LBound(configurations, 1) To UBound(configurations, 1)
        If UCase(conf_name) = UCase(configurations(i, 1)) Then
            Value = configurations(i, 2)
        End If
    Next
End Property

Public Sub ResetConfig(Optional void = Empty)
    If existConfigSheet Then
        With ThisWorkbook.Sheets("Config")
            Application.DisplayAlerts = False
            .Visible = xlSheetHidden
            .Delete
            Application.DisplayAlerts = True
        End With
    End If
    
    Dim configSheet As Worksheet
    Set configSheet = ThisWorkbook.Worksheets.Add(Sheets(1))
    configSheet.Name = "Config"
    Dim i, j
    Dim c As Collection: Set c = InitialSettings
    For i = 1 To c.Count
        Dim arr: arr = c(i)
        For j = LBound(arr) To UBound(arr)
            configSheet.Cells(i, j + 1).Value = arr(j)
        Next
    Next
    configSheet.Cells.EntireColumn.AutoFit
End Sub

Private Function existConfigSheet() As Boolean
    Dim ret As Boolean
    Dim sh As Object
    For Each sh In ThisWorkbook.Sheets
        If UCase(sh.Name) = UCase("Config") Then
            existConfigSheet = True
            Exit Function
        End If
    Next
End Function
