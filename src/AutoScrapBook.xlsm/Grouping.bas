Attribute VB_Name = "Grouping"
Public Sub GroupOverlappingShape(ByVal target_sheet As Worksheet)
    Dim SW() As Variant
    Dim C As Collection: Set C = WrappedShapes(target_sheet)
    For i = 1 To C.Count
        SW = C(i)
        If UBound(SW) > 0 Then
            target_sheet.Shapes.Range(SW).Group
        End If
    Next
End Sub

Private Function WrappedShapes(ByVal target_sheet As Worksheet) As Collection
    'シェイプをShapeWrapperで包んでコレクションに追加
    Dim C As New Collection, s As Shape, SW1 As ShapeWrapper
    For Each s In ActiveSheet.Shapes
        Set SW1 = New ShapeWrapper
        SW1.SetShape s
        C.Add SW1, SW1.Name
    Next

    'コレクションの各シェイプ同士の重なり判定
    Dim c2 As Collection: Set c2 = New Collection
    
    Dim SW2 As ShapeWrapper
    For Each SW1 In C
        Dim arr() As Variant
        ReDim arr(0)
        arr(0) = SW1.Name
        C.Remove SW1.Name
        For Each SW2 In C
            If Not (SW1 Is SW2) Then
                If SW1.IsOverlapped(SW2) Then
                    ReDim Preserve arr(UBound(arr) + 1)
                    arr(UBound(arr)) = SW2.Name
                    C.Remove SW2.Name
                End If
            End If
        Next
        c2.Add arr
    Next
    Set WrappedShapes = c2
End Function

Public Sub UngroupAllShapes(ByVal target_sheet As Worksheet)
    Dim sh As Shape
    For Each sh In ActiveSheet.Shapes
        Call RecUngroupShape(sh)
    Next
End Sub

Private Sub RecUngroupShape(sh As Shape)
    Dim memberShape As Shape
    If sh.Type = msoGroup Then
        For Each memberShape In sh.Ungroup
            Call RecUngroupShape(memberShape)
        Next
    End If
End Sub

