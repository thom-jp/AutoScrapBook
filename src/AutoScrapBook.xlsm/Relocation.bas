Attribute VB_Name = "Relocation"
Sub Main()
    Dim x As ILocator
    For Each x In GetRangeLocators(Sheet1)
        'Debug.Print x.Top
    Next
    
    Grouping.UngroupAllShapes Sheet1
    Grouping.GroupOverlappingShape Sheet1
    For Each x2 In GetShapeLocators(Sheet1)
        Debug.Print x2.Top
    Next
End Sub

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

Function LocateKey(L As ILocator) As Double
    LocateKey = L.Top
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
