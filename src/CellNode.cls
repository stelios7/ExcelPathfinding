VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private g_cost As Integer
Private h_cost As Integer
Private f_cost As Integer
Private vCell As Range
Private vParent As CellNode

Public Property Get Parent() As CellNode
    Set Parent = vParent
End Property

Public Property Set Parent(value As CellNode)
    Set vParent = value
    g_cost = value.GCost + 1
End Property

Public Property Get GCost() As Integer
    GCost = g_cost
End Property

Public Property Get HCost() As Integer
    HCost = h_cost
End Property

Public Property Get FCost() As Integer
    FCost = g_cost + h_cost
End Property

Public Property Get Cell() As Range
    Set Cell = vCell
End Property

Public Property Set Cell(value As Range)
    Set vCell = value
    
    'THESE ARE THE ROW AND COL OF THE TARGET CELL
    Dim i As Integer, j As Integer
    i = Range(STARTING_CELL_ADDRESS).Row + PUZZLE_SIZE - 1
    j = Range(STARTING_CELL_ADDRESS).Column + PUZZLE_SIZE - 1
    
    h_cost = vCell.Row - i + vCell.Column - j
    h_cost = Math.Abs(h_cost)
    
End Property

Public Property Get IsValid() As Boolean
    If Me.Cell.Row >= 2 And _
        Me.Cell.Column >= 2 And _
        Not Me.Cell.Address = STARTING_CELL_ADDRESS And _
        Not Me.Cell.Interior.Color = 0 Then
        IsValid = True
    Else
        IsValid = False
    End If
End Property


