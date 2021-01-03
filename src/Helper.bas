Attribute VB_Name = "Helper"
Option Explicit

Public Sub DebugNode(vNode As CellNode)
    Dim text As String
    text = "Current Cell: " & vNode.Cell.Address & _
        " G Cost: " & vNode.GCost & _
        " H Cost: " & vNode.HCost & _
        " F Cost: " & vNode.FCost
    If vNode.Parent Is Nothing Then
        text = "No Parent" & vbNewLine & text
    Else
        text = "Parent: " & vNode.Parent.Cell.Address & vbNewLine & text
    End If
    Debug.Print text
End Sub

