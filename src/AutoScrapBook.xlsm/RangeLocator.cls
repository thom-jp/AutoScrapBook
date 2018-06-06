VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeLocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ILocator
Private innerRange As Range

Public Sub SetRange(r As Range)
    Set innerRange = r
End Sub

Private Property Get ILocator_Top() As Long
    ILocator_Top = innerRange(1).Row
End Property

Private Property Get ILocator_Left() As Long
    ILocator_Left = innerRange(1).Column
End Property

Private Property Get ILocator_Bottom() As Long
    ILocator_Bottom = ILocator_Top + innerRange.Rows.Count - 1
End Property

Private Property Get ILocator_Right() As Long
    ILocator_Right = ILocator_Left + innerRange.Columns.Count - 1
End Property

Private Sub ILocator_Locate(r As Range)
    Dim o As Object: Set o = Selection
    innerRange.Copy
    r.PasteSpecial Paste:=xlPasteAllExceptBorders, Operation:=xlNone
    Application.CutCopyMode = False
    innerRange.ClearContents
    Set innerRange = r.Resize(innerRange.Rows.Count, innerRange.Columns.Count)
    o.Select
End Sub
