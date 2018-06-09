VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParagraphItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private InnerItem
Private InnerTop As Double

Property Get Top() As Double
    Top = InnerTop
End Property

Property Get Self() As ParagraphItem
    Set Self = Me
End Property

Property Get PrintableItem() As String
    If TypeName(Item) = "String" Then
        PrintableItem = Item
    Else
        PrintableItem = Item.Name
    End If
End Property

Property Set Item(obj As Object)
    InnerTop = obj.Top
    If TypeName(obj) = "Range" Then
        If IsNumeric(obj.Value) Then
            Let InnerItem = Format(obj.Value, obj.NumberFormat)
        Else
            Let InnerItem = obj.Value
        End If
    ElseIf TypeName(obj) = "Shape" Then
        Set InnerItem = obj
    Else
        Err.Raise 1000, , "ParagraphItemÇÕRangeÇ∆ShapeÇÃÇ›äiî[Ç≈Ç´Ç‹Ç∑ÅB"
    End If
End Property

Property Get Item() As Variant
    If IsObject(InnerItem) Then
        Set Item = InnerItem
    Else
        Let Item = InnerItem
    End If
End Property
