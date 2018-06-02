Attribute VB_Name = "Grouping"
Public Type Node
    x As Single
    y As Single
End Type

Private Function WrappedShapes() As Collection
    'シェイプをShapeWrapperで包んでコレクションに追加
    Dim c As New Collection, s As Shape, SW1 As ShapeWrapper
    For Each s In ActiveSheet.Shapes
        Set SW1 = New ShapeWrapper
        SW1.SetShape s
        c.Add SW1, SW1.Name
    Next

    'コレクションの各シェイプ同士の重なり判定
    Dim c2 As Collection: Set c2 = New Collection
    
    Dim SW2 As ShapeWrapper
    For Each SW1 In c
        Dim arr() As Variant
        ReDim arr(0)
        arr(0) = SW1.Name
        c.Remove SW1.Name
        For Each SW2 In c
            If Not (SW1 Is SW2) Then
                If SW1.IsOverlapped(SW2) Then
                    ReDim Preserve arr(UBound(arr) + 1)
                    arr(UBound(arr)) = SW2.Name
                    c.Remove SW2.Name
                End If
            End If
        Next
        c2.Add arr
    Next
    Set WrappedShapes = c2
End Function

Private Sub RecUngroupShape(sh As Shape)
    Dim memberShape As Shape
    If sh.Type = msoGroup Then
        For Each memberShape In sh.Ungroup
            Call RecUngroupShape(memberShape)
        Next
    End If
End Sub

Public Sub GroupOverlappingShape()
    Dim SW() As Variant
    Dim c As Collection: Set c = WrappedShapes
    For i = 1 To c.Count
        SW = c(i)
        ActiveSheet.Shapes.Range(SW).Group
    Next
End Sub

Public Sub UngroupAllShapes()
    Dim sh As Shape
    For Each sh In ActiveSheet.Shapes
        Call RecUngroupShape(sh)
    Next
End Sub
