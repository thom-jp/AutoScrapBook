VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private InnerShape As Shape

Public Property Get Name() As String
    Name = InnerShape.Name
End Property

Public Sub SetShape(s As Shape)
    Set InnerShape = s
End Sub

Public Property Get Top() As Single
    Top = InnerShape.Top
End Property

Public Property Get Left() As Single
    Left = InnerShape.Left
End Property

Public Property Get Bottom() As Single
    Bottom = InnerShape.Top + InnerShape.Height
End Property

Public Property Get Right() As Single
    Right = InnerShape.Left + InnerShape.Width
End Property

Public Property Get Nodes(Number As Integer) As Node
    Select Case Number
        Case 1
            Nodes.x = Me.Left
            Nodes.y = Me.Top
        Case 2
            Nodes.x = Me.Right
            Nodes.y = Me.Top
        Case 3
            Nodes.x = Me.Right
            Nodes.y = Me.Bottom
        Case 4
            Nodes.x = Me.Left
            Nodes.y = Me.Bottom
        Case Else
            Err.Raise 1000, , "1~4‚ðŽw’è‚µ‚Ä‚­‚¾‚³‚¢B"
    End Select
End Property

Public Function IsOverlapped(SW As ShapeWrapper) As Boolean
    Dim i As Integer
    For i = 1 To 4 Step 1
        IsOverlapped = _
            (SW.Nodes(i).x > Me.Left And _
            SW.Nodes(i).x < Me.Right And _
            SW.Nodes(i).y > Me.Top And _
            SW.Nodes(i).y < Me.Bottom) _
            Or _
            (Me.Nodes(i).x > SW.Left And _
            Me.Nodes(i).x < SW.Right And _
            Me.Nodes(i).y > SW.Top And _
            Me.Nodes(i).y < SW.Bottom)
        If IsOverlapped Then Exit Function
    Next
End Function

