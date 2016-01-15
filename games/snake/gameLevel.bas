Attribute VB_Name = "gameLevel"
Option Explicit

Public Tick As Double
Public cellSize As Integer
Public boardHeight As Integer, boardWidth As Integer 'max 30 450mm
Public foodMatrix() As Integer
Public wallmatrix() As Integer
Public snake() As Integer
Dim wallStr As String
Dim foodStr As String
Dim startBodySize As Integer
Dim startRow As Integer, startColumn As Integer

Function levelCount() As Integer
    levelCount = 3
End Function

Sub levelChange(i As Integer)
    Select Case i
        Case 1
            level1
        Case 2
            level2
        Case 3
            level3
    End Select
End Sub

Private Sub level3()
    Dim i As Integer, e As Integer
    Dim x As Integer, y As Integer
    Dim foodSplit() As String
    Dim wallSplit() As String
    
    '/ level settings
    Tick = 0.3
    
    cellSize = 15
    boardHeight = 20
    boardWidth = 20
    
    startBodySize = 5
    startRow = 4
    startColumn = 10
    
    foodStr = "5,5,9,9"
    wallStr = "9,0,8,0,18,0,7,0,17,0,6,0,16,0,5,0,15,0,4,0,14,0,3,0,13,0,2,0,12,0,1,0,11,0,10,5,10,4,10,3,10,2,10,1,10,0,9,19,19,0,19,1,19,2,19,3,15,4,16,4,17,4,18,4,19,4,19,5,19,6,19,7,19,8,19,9,19,10,19,11,19,12,19,13,19,14,19,15,19,16,7,17,8,17,9,17,10,17,11,17,12,17,13,17,14,17,15,17,16,17,17,17,18,17,19,17,19,18,19,19,8,19,18,19,7,19,17,19,6,19,16,19,5,19,15,19,4,19,14,19,3,19,13,19,2,19,12,19,1,19,11,19,0,0,0,1,0,2,0,3,0,4,0,5,0,6,0,7,6,8,5,8,4,8,3,8,2,8,1,8,0,8,0,9,0,10,0,11,0,12,0,13,0,14,0,15,0,16,0,17,0,18,0,19,10,19"
    
    foodSplit = Split(foodStr, ",")
    wallSplit = Split(wallStr, ",")
    '/ end settings
    
    ReDim foodMatrix(boardWidth - 1, boardHeight - 1)
    ReDim wallmatrix(boardWidth - 1, boardHeight - 1)
    
    For i = 0 To UBound(foodSplit)
        foodMatrix(foodSplit(i), foodSplit(i + 1)) = 1
        i = i + 1
    Next i
    
    For i = 0 To UBound(wallSplit)
        wallmatrix(wallSplit(i), wallSplit(i + 1)) = 1
        i = i + 1
    Next i
    
    ReDim snake(1, startBodySize - 1)
    For i = 0 To startBodySize - 1
        snake(0, i) = startColumn - 1 - i
        snake(1, i) = startRow - 1
    Next i
End Sub

Private Sub level2()
    Dim i As Integer, e As Integer
    Dim x As Integer, y As Integer
    Dim foodSplit() As String
    Dim wallSplit() As String
    
    '/ level settings
    Tick = 0.3
    
    cellSize = 15
    boardHeight = 15
    boardWidth = 15
    
    startBodySize = 5
    startRow = 4
    startColumn = 10
    
    foodStr = "5,5,10,9"
    wallStr = "14,11,13,12,14,12,12,13,13,13,14,13,11,14,12,14,13,14,14,14,14,3,13,2,14,2,12,1,13,1,14,1,11,0,12,0,13,0,14,0,0,3,1,2,0,2,2,1,1,1,0,1,3,0,2,0,1,0,0,0,0,11,1,12,0,12,6,7,8,7,7,6,7,8,7,7,2,13,1,13,0,13,3,14,2,14,1,14,0,14"
    
    foodSplit = Split(foodStr, ",")
    wallSplit = Split(wallStr, ",")
    '/ end settings
    
    ReDim foodMatrix(boardWidth - 1, boardHeight - 1)
    ReDim wallmatrix(boardWidth - 1, boardHeight - 1)
    
    For i = 0 To UBound(foodSplit)
        foodMatrix(foodSplit(i), foodSplit(i + 1)) = 1
        i = i + 1
    Next i
    
    For i = 0 To UBound(wallSplit)
        wallmatrix(wallSplit(i), wallSplit(i + 1)) = 1
        i = i + 1
    Next i
    
    ReDim snake(1, startBodySize - 1)
    For i = 0 To startBodySize - 1
        snake(0, i) = startColumn - 1 - i
        snake(1, i) = startRow - 1
    Next i
End Sub

Private Sub level1()
    Dim i As Integer, e As Integer
    Dim x As Integer, y As Integer
    Dim foodSplit() As String
    Dim wallSplit() As String
    
    '/ level settings
    Tick = 0.3
    
    cellSize = 15
    boardHeight = 20
    boardWidth = 10
    
    startBodySize = 5
    startRow = 4
    startColumn = 6
    
    foodStr = "5,5,7,7"
    wallStr = "9,15,9,16,9,17,9,18,9,19,9,10,8,10,0,10,1,10,0,4,0,3,0,2,0,1,0,0,6,4,6,3,6,2,6,1,6,0,3,15,3,16,3,17,3,18,3,19"
    
    foodSplit = Split(foodStr, ",")
    wallSplit = Split(wallStr, ",")
    '/ end settings
    
    ReDim foodMatrix(boardWidth - 1, boardHeight - 1)
    ReDim wallmatrix(boardWidth - 1, boardHeight - 1)
    
    For i = 0 To UBound(foodSplit)
        foodMatrix(foodSplit(i), foodSplit(i + 1)) = 1
        i = i + 1
    Next i
    
    For i = 0 To UBound(wallSplit)
        wallmatrix(wallSplit(i), wallSplit(i + 1)) = 1
        i = i + 1
    Next i
    
    ReDim snake(1, startBodySize - 1)
    For i = 0 To startBodySize - 1
        snake(0, i) = startColumn - 1 - i
        snake(1, i) = startRow - 1
    Next i
End Sub
