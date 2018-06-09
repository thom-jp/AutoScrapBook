VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeLocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ILocator
Private InnerShape As Shape

Public Property Get Name() As String
    Name = InnerShape.Name
End Property

Public Property Get Self() As Object
    Set Self = Me
End Property

Public Sub SetShape(s As Shape)
    Set InnerShape = s
End Sub

Public Property Get Top() As Long
    Top = ILocator_Top
End Property

Public Property Get Left() As Long
    Left = ILocator_Left
End Property

Public Property Get Bottom() As Long
    Bottom = ILocator_Bottom
End Property

Public Property Get Right() As Long
    Right = ILocator_Right
End Property

Private Property Get ILocator_Top() As Long
    Dim r As Range: Set r = InnerShape.Parent.Range("a1")
    Do While InnerShape.Top > r.Top
        Set r = r.Offset(1, 0)
    Loop
    ILocator_Top = r.Row
End Property

Private Property Get ILocator_Bottom() As Long
    Dim r As Range: Set r = InnerShape.Parent.Range("a1")
    Do While InnerShape.Top + InnerShape.Height > r.Top
        Set r = r.Offset(1, 0)
    Loop
    ILocator_Bottom = r.Row - 1
End Property

Private Property Get ILocator_Left() As Long
    Dim r As Range: Set r = InnerShape.Parent.Range("a1")
    Do While InnerShape.Left > r.Left
        Set r = r.Offset(0, 1)
    Loop
    ILocator_Left = r.Column
End Property

Private Property Get ILocator_Right() As Long
    Dim r As Range: Set r = InnerShape.Parent.Range("a1")
    Do While InnerShape.Left + InnerShape.Width > r.Left
        Set r = r.Offset(0, 1)
    Loop
    ILocator_Right = r.Column - 1
End Property

Private Sub ILocator_Locate(r As Range)
    InnerShape.Top = r.Top
    InnerShape.Left = r.Left
End Sub
