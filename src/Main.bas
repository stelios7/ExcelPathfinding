Attribute VB_Name = "Main"
Option Explicit

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
Sub SolvePuzzle()
    'TO DO
    Application.Calculation = xlCalculationManual
    Dim myTimer As Double
    myTimer = Timer
    Dim customColor As Variant
    customColor = RGB(Rnd() * 255, Rnd() * 255, Rnd() * 255)
    
    Dim d_Openionary As Object
    Set d_Openionary = CreateObject("Scripting.Dictionary")
    
    Dim startingCell As Range, finishCell As Range
    Set startingCell = ActiveSheet.Range("B2")
    Set finishCell = Cells(startingCell.Row + PUZZLE_HEIGHT - 1, startingCell.Column + PUZZLE_WIDTH - 1)
    
    Dim puzzleField As Range
    Set puzzleField = ActiveSheet.Range(startingCell, finishCell)
    
    Dim d_Open As Object
    Set d_Open = CreateObject("Scripting.Dictionary")
    
    Dim d_Searched As Object
    Set d_Searched = CreateObject("Scripting.dictionary")
    
    Dim startingNode As CellNode
    Set startingNode = New CellNode
    Set startingNode.Cell = ActiveSheet.Range(STARTING_CELL_ADDRESS)
    
    
    
    d_Open.Add STARTING_CELL_ADDRESS, startingNode
    
    'VARIABLES FOR CLOSEST NODE, DUMMY NODE ADJACENT d_open AND A COLLECTION TO HOLD THEM
    Dim closestNode As CellNode, vNode As CellNode, adjacent As CellNode
    Dim adjacentList As Collection
    
    'SOME CUSTOM VARIABLES TO AVOID USING WHOLE OBJECTS AGAIN AND AGAIN
    Dim nodeCost As Integer, nHCost As Integer, clNodeFCost As Integer, clNodeHCost As Integer
    Dim success As Boolean
    Dim currentNodeAddress As String
    
    Do While d_Open.Count > 0
        Dim key As Variant
        
        For Each key In d_Open.keys
            Set closestNode = d_Open(key)
            Exit For
        Next
        
        For Each key In d_Open.keys
            With d_Open(key)
                With .Cell
                    currentNodeAddress = .Address
                End With
                nodeCost = .FCost
                nHCost = .HCost
                clNodeFCost = closestNode.FCost
                clNodeHCost = closestNode.HCost
                If Not currentNodeAddress = closestNode.Cell.Address Then
                    If nodeCost <= clNodeFCost Then
                        Set closestNode = d_Open(key)
                    End If
                End If
            End With
        Next
        
        With closestNode.Cell
            Dim traverseColor As Long, maxDistance As Integer, colorOffset As Long
            maxDistance = startingNode.HCost
            colorOffset = RGB(0, closestNode.HCost / maxDistance * 150, 0)
            traverseColor = RGB(0, Rnd() * 20 + 50, Rnd() * 20 + 220) + colorOffset
            .Interior.color = traverseColor
            d_Open.Remove .Address
        End With
        
        Dim bCell As Range, rCell As Range, tCell As Range, lCell As Range
        Set adjacentList = New Collection
        
        'CHECK BOTTOM CELL
        Set bCell = closestNode.Cell.Offset(1, 0)
        Set adjacent = New CellNode
        Set adjacent.Cell = bCell
        Set adjacent.Parent = closestNode
        If adjacent.IsValid Then
            adjacentList.Add adjacent
        End If
        
        'CHECK RIGHT CELL
        Set rCell = closestNode.Cell.Offset(0, 1)
        Set adjacent = New CellNode
        Set adjacent.Cell = rCell
        Set adjacent.Parent = closestNode
        If adjacent.IsValid Then
            adjacentList.Add adjacent
        End If
        
        'CHECK TOP CELL
        Set tCell = closestNode.Cell.Offset(-1, 0)
        Set adjacent = New CellNode
        Set adjacent.Cell = tCell
        Set adjacent.Parent = closestNode
        If adjacent.IsValid Then
            adjacentList.Add adjacent
        End If
        
        'CHECK LEFT CELL
        Set lCell = closestNode.Cell.Offset(0, -1)
        Set adjacent = New CellNode
        Set adjacent.Cell = lCell
        Set adjacent.Parent = closestNode
        If adjacent.IsValid Then
            adjacentList.Add adjacent
        End If
        
        Dim v As CellNode
        For Each v In adjacentList
            
            With v.Cell
                currentNodeAddress = .Address
                
                If currentNodeAddress = finishCell.Address Then
                    Debug.Print "Stop Search"
                    success = True
                    Exit Do
                End If
                
                If d_Open.exists(currentNodeAddress) Then
                    If v.FCost < d_Open(currentNodeAddress).FCost Then
                        Set d_Open(currentNodeAddress) = v
                        Set d_Open(currentNodeAddress) = v
                    End If
                End If
                
                Dim text As String
                text = v.FCost & vbNewLine & v.GCost & "|" & v.HCost
                
                
                If Not d_Searched.exists(currentNodeAddress) Then
                    If Not d_Open.exists(currentNodeAddress) Then
                        .Interior.color = traverseColor + RGB(0, 50, 0)
                        d_Open.Add currentNodeAddress, v
                        .value = text
                    Else
                        If v.FCost < d_Open(currentNodeAddress).FCost Then
                            Set d_Open(currentNodeAddress) = v
                        End If
                    End If
                Else
                    
                    If v.GCost < d_Searched(currentNodeAddress).GCost Then
                        Set d_Searched(currentNodeAddress) = v
'                        .value = text
                    End If
                End If
                
            End With
        Next
        
        d_Searched.Add closestNode.Cell.Address, closestNode
'        closestNode.Cell.Interior.Color = RGB(0, 100, Rnd() * 20 + 200)

    Loop
    Dim totalTime As String
    totalTime = "Total time: " & Format(Timer - myTimer, ".00")
    
    If Not success Then
        Set vNode = New CellNode
        Application.ScreenUpdating = False
        For Each key In d_Searched.keys
'            d_Searched(key).Cell.Interior.color = RGB(Rnd() * 30 + 50, 10, 100)
        Next
        Debug.Print "Total time: " & totalTime
            
    Else
        Debug.Print "Steps: " & TraceBackFrom(v), totalTime
    End If
    
End Sub

Private Function TraceBackFrom(vNode As CellNode) As Integer
    Dim path As New Collection
    Do While Not vNode.Parent Is Nothing
        path.Add vNode.Cell
        Set vNode = vNode.Parent
    Loop
    
    Dim i As Integer
    Dim customColor As Variant
    customColor = RGB(255, Rnd() * 20 + 50, Rnd() * 20 + 200)
    For i = 1 To path.Count
        path(i).Interior.color = customColor
        Sleep BACK_TRACE_SPEED
    Next
    
    TraceBackFrom = path.Count
End Function

Sub GeneratePuzzle()
    Application.ScreenUpdating = False
    ActiveSheet.Cells.Clear
    Dim startingCell As Range
    Set startingCell = ActiveSheet.Range("B2")
    
    'SET LAST CELL
    Dim finishCell As Range
    Set finishCell = startingCell.Offset(0, PUZZLE_WIDTH - 1).Offset(PUZZLE_HEIGHT - 1, 0)
       
    Dim puzzleBorders As Range
    Set puzzleBorders = ActiveSheet.Range(startingCell.Offset(-1, -1), finishCell.Offset(1, 1))
    puzzleBorders.Interior.color = vbBlack
    
    Dim puzzleField As Range
    Set puzzleField = ActiveSheet.Range(startingCell, finishCell)
    puzzleField.Interior.color = vbWhite
    
    startingCell.Interior.color = vbBlue
    finishCell.Interior.color = vbRed
    
    'CREATE RANDOM BLOCKS INSIDE PUZZLEFIELD
    Dim vCell As Range
    Dim rng As Double
    For Each vCell In puzzleField
        If Not vCell.Address = STARTING_CELL_ADDRESS And Not vCell.Address = finishCell.Address Then
            rng = Rnd()
            If rng > 1 - BRICK_DENSITY Then
'                Sleep 1
                vCell.Interior.color = vbBlack
            End If
        End If
        
    Next
    
    With puzzleField.Cells
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 8
    End With
    
    Application.ScreenUpdating = True
End Sub

