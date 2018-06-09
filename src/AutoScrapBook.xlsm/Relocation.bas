Attribute VB_Name = "Relocation"
Sub Main()
    Call Config.LoadConfig
    RelocateAll Sheet1
End Sub

Sub RelocateAll(ByVal target_sheet As Worksheet)
    'テキストと画像を各Locatorにセットして混在コレクションを作り、ソートする。
    Dim c As Collection
    Grouping.UngroupAllShapes target_sheet
    Grouping.GroupOverlappingShape target_sheet
    Set c = MargeCollection(GetRangeLocators(target_sheet), GetShapeLocators(target_sheet))
    Call CollectionSort.CSort(c, "LocateKey")
    
    'テキストは上書き防止の為、一時退避処理
    With ActiveSheet
        Dim sh As Worksheet: Set sh = Worksheets.Add
        .Activate
    End With
    Dim loc As ILocator
    Dim n As Long: n = 1
    For Each loc In c
        If loc.LocatorType = eRangeLocator Then
            loc.Locate sh.Cells(n, 1)
            n = n + (loc.Bottom - loc.Top) + 2
        End If
    Next
    
    '再配置
    Dim r As Long: r = Config.startRow
    For Each loc In c
        If loc.LocatorType = eRangeLocator Then
            loc.Locate target_sheet.Cells(r, 1)
            r = loc.Bottom + 2
        Else
            loc.Locate target_sheet.Cells(r, Config.startColumn)
            r = loc.Bottom + Config.Margin
        End If
    Next
    
    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True
End Sub

Function LocateKey(L As ILocator) As Double
    LocateKey = L.Top
End Function

Function MargeCollection(c1 As Collection, c2 As Collection) As Collection
    Dim ret As Collection: Set ret = New Collection
    Dim x As Variant
    For Each x In c1
        ret.Add x
    Next
    For Each x In c2
        ret.Add x
    Next
    Set MargeCollection = ret
End Function

Function GetShapeLocators(target_sheet As Worksheet) As Collection
    Dim ret As Collection: Set ret = New Collection
    Dim sh As Shape
    For Each sh In target_sheet.Shapes
        With New ShapeLocator
            .SetShape sh
            ret.Add .Self
        End With
    Next
    Set GetShapeLocators = ret
End Function

Function GetRangeLocators(target_sheet As Worksheet) As Collection
    Dim ret As Collection: Set ret = New Collection
    
    With target_sheet.Cells.SpecialCells(xlCellTypeLastCell)
        Dim maxColumn As Long: maxColumn = .Column
        Dim maxRow As Long: maxRow = .Row
    End With
    
    With target_sheet
        Dim line As Range
        Set line = .Range(.Cells(1, 1), .Cells(1, maxColumn))
    End With
    
    Do
        Do While isBlank(line)
            Set line = line.Offset(1, 0)
            If line.Row > maxRow Then GoTo Fin
        Loop
        
        Dim block As Range
        Set block = line.Cells
        
        Do Until isBlank(line.Offset(1, 0))
            Set line = line.Offset(1, 0)
            Set block = Union(block, line)
        Loop
        Set line = line.Offset(1, 0)
        
        With New RangeLocator
            .SetRange block.Cells
            ret.Add .Self
        End With
    Loop
Fin:
    Set GetRangeLocators = ret
End Function

Private Function isBlank(target_range As Range) As Boolean
    Dim r As Range
    For Each r In target_range
        If r.Value <> "" Then GoTo DataFound
    Next
    isBlank = True
Exit Function
DataFound:
    isBlank = False
End Function
