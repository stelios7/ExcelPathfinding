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
    If Me.Cell.Address = STARTING_CELL_ADDRESS Then
        GCost = 0
        Exit Property
    End If
    
    If Me.Parent Is Nothing Then
        Debug.Print "Node hasn't been explored yet"
        GCost = 999
    Else
        GCost = g_cost
    End If
End Property

Public Property Get HCost() As Integer
    HCost = h_cost
End Property

Public Property Get FCost() As Integer
    If Me.Cell.Address = STARTING_CELL_ADDRESS Then
        FCost = h_cost
        Exit Property
    End If
    If Me.Parent Is Nothing Then
        Debug.Print "Node hasn't been explored yet"
        FCost = 999
    Else
        FCost = g_cost + h_cost
    End If
End Property

Public Property Get Cell() As Range
    Set Cell = vCell
End Property

Public Property Set Cell(value As Range)
    Set vCell = value
    
    'THESE ARE THE ROW AND COL OF THE TARGET CELL
    Dim i As Integer, j As Integer
    i = Range(STARTING_CELL_ADDRESS).Row + PUZZLE_HEIGHT - 1
    j = Range(STARTING_CELL_ADDRESS).Column + PUZZLE_WIDTH - 1
    
    h_cost = (vCell.Row - i) ^ 2 + (vCell.Column - j) ^ 2
    h_cost = h_cost
    
End Property

Public Property Get IsValid() As Boolean
    If Me.Cell.Row >= 2 And _
        Me.Cell.Column >= 2 And _
        Not Me.Cell.Address = STARTING_CELL_ADDRESS And _
        Not Me.Cell.Interior.color = 0 Then
        IsValid = True
    Else
        IsValid = False
    End If
End Property


