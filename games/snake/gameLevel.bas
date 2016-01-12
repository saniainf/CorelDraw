Attribute VB_Name = "gameLevel"
Option Explicit

Public Tick As Double
Public cellSize As Integer
Public boardHeight As Integer, boardWidth As Integer 'max 30 450mm
Public startBodySize As Integer
Public startRow As Integer, startColumn As Integer
Public foodCount As Integer
Public foodMatrix() As Integer
Public snake() As Integer

Function levelCount() As Integer
    levelCount = 2
End Function

Sub levelChange(i As Integer)
    Select Case i
        Case 1
            level1
        Case 2
            level2
    End Select
End Sub

Private Sub level1()
    Dim i As Integer, e As Integer
    Dim X As Integer, Y As Integer
    Tick = 0.3
    
    cellSize = 15
    boardHeight = 15
    boardWidth = 15
    
    startBodySize = 5
    startRow = 4
    startColumn = 5
    
    foodCount = 2
    
    ReDim snake(1, startBodySize - 1)
    For i = 0 To startBodySize - 1
        snake(0, i) = startColumn - 1 - i
        snake(1, i) = startRow - 1
    Next i
    
    ReDim foodMatrix(boardWidth - 1, boardHeight - 1)
    Randomize
    For e = 0 To foodCount - 1
        X = Int(boardWidth * Rnd)
        Y = Int(boardHeight * Rnd)
        foodMatrix(X, Y) = 1
    Next e
End Sub

Private Sub level2()
    Dim i As Integer, e As Integer
    Dim X As Integer, Y As Integer
    Tick = 0.3
    
    cellSize = 15
    boardHeight = 25
    boardWidth = 15
    
    startBodySize = 5
    startRow = 4
    startColumn = 5
    
    foodCount = 4
    
    ReDim snake(1, startBodySize - 1)
    For i = 0 To startBodySize - 1
        snake(0, i) = startColumn - 1 - i
        snake(1, i) = startRow - 1
    Next i
    
    ReDim foodMatrix(boardWidth - 1, boardHeight - 1)
    Randomize
    For e = 0 To foodCount - 1
        X = Int(boardWidth * Rnd)
        Y = Int(boardHeight * Rnd)
        foodMatrix(X, Y) = 1
    Next e
End Sub
