Attribute VB_Name = "Relocation"
Sub Main()
    Dim x As ILocator
    For Each x In GetRangeLocators(Sheet1)
        Debug.Print x.Top
    Next
End Sub

Function GetRangeLocators(targetSheet As Worksheet) As Collection
    Dim ret As Collection: Set ret = New Collection
    
    With targetSheet.Cells.SpecialCells(xlCellTypeLastCell)
        Dim maxColumn As Long: maxColumn = .Column
        Dim maxRow As Long: maxRow = .Row
    End With
    
    With targetSheet
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

Private Function isBlank(targetRange As Range) As Boolean
    Dim r As Range
    For Each r In targetRange
        If r.Value <> "" Then GoTo DataFound
    Next
    isBlank = True
Exit Function
DataFound:
    isBlank = False
End Function
