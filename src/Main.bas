Attribute VB_Name = "Main"
Option Explicit

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
Sub SolvePuzzle()
    'TO DO
    Dim customColor As Variant
    customColor = RGB(Rnd() * 255, Rnd() * 255, Rnd() * 255)
    
    Dim nodes As Object
    Set nodes = CreateObject("Scripting.Dictionary")
    
    Dim openDictionary As Object
    Set openDictionary = CreateObject("Scripting.Dictionary")
    
    Dim startingCell As Range, finishCell As Range
    Set startingCell = ActiveSheet.Range("B2")
    Set finishCell = Cells(startingCell.Row + PUZZLE_SIZE - 1, startingCell.Column + PUZZLE_SIZE - 1)
    
    Dim puzzleField As Range
    Set puzzleField = ActiveSheet.Range(startingCell, finishCell)
    
    Dim vCell As Range
    Dim myNode As CellNode
    For Each vCell In puzzleField
        Set myNode = New CellNode
        Set myNode.Cell = vCell
        nodes.Add vCell.Address, myNode
    Next
    
    Dim openList As Object
    Set openList = CreateObject("Scripting.Dictionary")
    
    Dim closedList As Object
    Set closedList = CreateObject("Scripting.dictionary")
    
    openList.Add STARTING_CELL_ADDRESS, nodes(STARTING_CELL_ADDRESS)
    Dim closestNode As CellNode, vNode As CellNode
    Dim successor As CellNode, successorList As Collection
    Dim nodeCost As Integer, nHCost As Integer, closestNodeCost As Integer
    
    Do While openList.Count > 0
        Dim key As Variant
        For Each key In openList.keys
            Set closestNode = openList(key)
'            Exit For
        Next
        For Each key In openList.keys
        
            nodeCost = openList(key).FCost
            nHCost = openList(key).HCost
            closestNodeCost = closestNode.FCost
            'Debug.Print key, openList(key).GCost, nodeCost, closestNodeCost
            If nodeCost < closestNodeCost Then
                Set closestNode = openList(key)
                Exit For
            End If
            

'            If nHCost < closestNode.HCost Then
'                Set closestNode = openList(key)
'            End If
            
        Next
        openList.Remove closestNode.Cell.Address
        
        
        Dim bCell As Range, rCell As Range, tCell As Range, lCell As Range
        Set successorList = New Collection
        'CHECK BOTTOM CELL
        Set bCell = closestNode.Cell.Offset(1, 0)
        Set successor = New CellNode
        Set successor.Cell = bCell
        Set successor.Parent = closestNode
        If successor.IsValid And Not closedList.exists(successor.Cell.Address) Then
            successorList.Add successor
'            Debug.Print vbTab & successor.Cell.Address & " added to the list"
        End If
'        Helper.DebugNode successor
        
        'CHECK RIGHT CELL
        Set rCell = closestNode.Cell.Offset(0, 1)
        Set successor = New CellNode
        Set successor.Cell = rCell
        Set successor.Parent = closestNode
'        Debug.Print successor.Parent Is Nothing
        If successor.IsValid And Not closedList.exists(successor.Cell.Address) Then
            successorList.Add successor
'            Debug.Print vbTab & successor.Cell.Address & " added to the list"
        End If
'        Helper.DebugNode successor
        
        'CHECK TOP CELL
        Set tCell = closestNode.Cell.Offset(-1, 0)
        Set successor = New CellNode
        Set successor.Cell = tCell
        Set successor.Parent = closestNode
        If successor.IsValid And Not closedList.exists(successor.Cell.Address) Then
            successorList.Add successor
'            Debug.Print vbTab & successor.Cell.Address & " added to the list"
        End If
        
        'CHECK LEFT CELL
        Set lCell = closestNode.Cell.Offset(0, -1)
        Set successor = New CellNode
        Set successor.Cell = lCell
        Set successor.Parent = closestNode
        If successor.IsValid And Not closedList.exists(successor.Cell.Address) Then
            successorList.Add successor
'            Debug.Print vbTab & successor.Cell.Address & " added to the list"
        End If
        
        Dim v As CellNode
        For Each v In successorList
'            Set nodes(v.Cell.Address).Parent = v.Parent
            If v.Cell.Address = finishCell.Address Then
                Debug.Print "Stop Search"
                TraceBackFrom v
                Exit Do
            End If
                
            Dim currentAddress As String
            currentAddress = v.Cell.Address
            
            If openList.exists(currentAddress) Then
'                If v.FCost < openList(currentaddress).FCost Then
'                    Set nodes(currentaddress) = v
'                End If
                If v.HCost < openList(currentAddress).HCost Then
                    Set nodes(currentAddress) = v
                End If
            End If
            

            If Not closedList.exists(currentAddress) Then
                If openList.exists(currentAddress) Then
'                    Helper.DebugNode v
'                    Helper.DebugNode openList(currentAddress)
                Else
                    openList.Add v.Cell.Address, v
'                    v.Cell.Interior.Color = customColor
                    v.Cell.Interior.Color = RGB(Rnd() * 10 + 10, Rnd() * 25 + 135, Rnd() * 25 + 220)
                End If
                
'                Sleep 1
            End If
        Next
        closedList.Add closestNode.Cell.Address, closestNode
    Loop
End Sub

Private Sub TraceBackFrom(vNode As CellNode)
    Dim path As New Collection
    Do While Not vNode.Parent Is Nothing
        path.Add vNode.Cell
        Set vNode = vNode.Parent
    Loop
    
    Dim i As Integer
    Dim customColor As Variant
    customColor = RGB(Rnd() * 255, Rnd() * 255, Rnd() * 255)
    For i = 1 To path.Count
        path(i).Interior.Color = customColor
        Sleep 10
    Next
    
    Debug.Print "Steps: " & path.Count
End Sub

Sub GeneratePuzzle()
    Application.ScreenUpdating = False
    Dim startingCell As Range
    Set startingCell = ActiveSheet.Range("B2")
    
    'SET LAST CELL
    Dim finishCell As Range
    Set finishCell = startingCell.Offset(0, PUZZLE_SIZE - 1).Offset(PUZZLE_SIZE - 1, 0)
       
    Dim puzzleBorders As Range
    Set puzzleBorders = ActiveSheet.Range(startingCell.Offset(-1, -1), finishCell.Offset(1, 1))
    puzzleBorders.Interior.Color = vbBlack
    
    Dim puzzleField As Range
    Set puzzleField = ActiveSheet.Range(startingCell, finishCell)
    puzzleField.Interior.Color = vbWhite
    
    startingCell.Interior.Color = vbBlue
    finishCell.Interior.Color = vbRed
    
    'CREATE RANDOM BLOCKS INSIDE PUZZLEFIELD
    Dim vCell As Range
    Dim rng As Double
    For Each vCell In puzzleField
        If Not vCell.Address = STARTING_CELL_ADDRESS And Not vCell.Address = finishCell.Address Then
            rng = Rnd()
            If rng > 0.65 Then
'                Sleep 1
                vCell.Interior.Color = vbBlack
            End If
        End If
        
    Next
    Application.ScreenUpdating = True
End Sub
