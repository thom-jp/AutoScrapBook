VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeLocator"
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

Public Property Get Top() As Long
    Dim r As Range: Set r = InnerShape.Parent.Range("a1")
    Do While InnerShape.Top > r.Top
        Set r = r.Offset(1, 0)
    Loop
    Top = r.Row
End Property

Public Property Get Left() As Long
    Dim r As Range: Set r = InnerShape.Parent.Range("a1")
    Do While InnerShape.Left > r.Left
        Set r = r.Offset(0, 1)
    Loop
    Left = r.Column
End Property

Public Property Get Bottom() As Long
    Dim r As Range: Set r = InnerShape.Parent.Range("a1")
    Do While InnerShape.Top + InnerShape.Height > r.Top
        Set r = r.Offset(1, 0)
    Loop
    Bottom = r.Row - 1
End Property

Public Property Get Right() As Long
    Dim r As Range: Set r = InnerShape.Parent.Range("a1")
    Do While InnerShape.Left + InnerShape.Width > r.Left
        Set r = r.Offset(0, 1)
    Loop
    Right = r.Column - 1
End Property

