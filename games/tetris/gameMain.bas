Attribute VB_Name = "gameMain"
Option Explicit

Dim imDone As Boolean
Dim fallTmr As Double
Dim cellSize As Integer
Dim boardHeight As Integer, boardWidth As Integer
Dim screenWidth As Integer, screenHeight As Integer
Dim fallTick As Double
Dim keyUp As Boolean, keyDown As Boolean, keyLeft As Boolean, keyRight As Boolean, keySpace As Boolean
Dim returnValue As String
Dim fMatrix() As Integer
Dim boardMatrix() As Integer
Dim fX As Integer, fY As Integer
Dim boardDraw As Boolean, figureDraw As Boolean
Dim cFigures As New Collection
Dim figO As New Collection, figI As New Collection, figS As New Collection, figZ As New Collection, figL As New Collection, figJ As New Collection, figT As New Collection
Dim currentFigure As Integer, currentShape As Integer

Sub LoadLevel()
    fallTick = 0.3
    cellSize = 15
    boardHeight = 20
    boardWidth = 10
End Sub

Function GameLoop() As String
    Initialization
    Do Until imDone
        DoEvents
        UpdateInput
        Update
        Draw
    Loop
    Destroy
    GameLoop = returnValue
End Function

Private Sub Initialization()
    imDone = False
    boardDraw = False
    figureDraw = False
    ReDim boardMatrix(boardWidth - 1, boardHeight - 1)
    drawGameField
    
    figO.Add strToArr("0,0,0,0,0,1,1,0,0,1,1,0,0,0,0,0")
    cFigures.Add figO
    
    figI.Add strToArr("0,0,0,0,1,1,1,1,0,0,0,0,0,0,0,0")
    figI.Add strToArr("0,0,1,0,0,0,1,0,0,0,1,0,0,0,1,0")
    cFigures.Add figI
    
    figS.Add strToArr("0,0,0,0,0,0,1,1,0,1,1,0,0,0,0,0")
    figS.Add strToArr("0,0,1,0,0,0,1,1,0,0,0,1,0,0,0,0")
    cFigures.Add figS
    
    nextFigure
End Sub

Private Sub nextFigure()
    Randomize
    currentFigure = Int((cFigures.Count * Rnd) + 1)
    currentShape = 1
    fMatrix = cFigures.Item(currentFigure).Item(currentShape)
    
    '/ figure start position
    fallTmr = Timer
    fX = 3
    fY = UBound(boardMatrix, 2) - UBound(fMatrix, 2)
    
    boardDraw = True
End Sub

Private Sub rotateFigure()
    Dim canRotate As Boolean
    Dim prevShape As Integer
    canRotate = True
    
    prevShape = currentShape
    currentShape = currentShape + 1
    If currentShape > cFigures.Item(currentFigure).Count Then
        currentShape = 1
    End If
    fMatrix = cFigures.Item(currentFigure).Item(currentShape)
    
    If collWallWell Then
        canRotate = False
    ElseIf collBotWell Then
        canRotate = False
    ElseIf collAnotherFigure Then
        canRotate = False
    End If
    
    If Not canRotate Then
        currentShape = prevShape
        fMatrix = cFigures.Item(currentFigure).Item(currentShape)
    End If
    figureDraw = True
End Sub

Private Function strToArr(s As String) As Integer()
    Dim e As Integer, i As Integer, n As Integer
    Dim aStr() As String
    Dim tmpMatrix(3, 3) As Integer
    aStr = Split(s, ",")
    n = 0
    For i = 0 To UBound(tmpMatrix, 2)
        For e = 0 To UBound(tmpMatrix, 1)
            tmpMatrix(e, i) = aStr(n)
            n = n + 1
        Next e
    Next i
    strToArr = tmpMatrix
End Function

Private Sub drawGameField()
    Dim s As Shape
    Dim e As Integer, i As Integer
    Application.Optimization = True
    For i = 0 To boardHeight - 1
        For e = 0 To boardWidth - 1
            Set s = ActivePage.Layers.Item(5).CreateRectangle(e * cellSize, i * cellSize, e * cellSize + cellSize, i * cellSize + cellSize)
            s.Fill.ApplyNoFill
            s.Outline.Color.CMYKAssign 0, 0, 0, 20
            s.Outline.width = 0.1
        Next e
    Next i
    ActiveDocument.ClearSelection
End Sub

Private Sub UpdateInput()
    If (GetAsyncKeyState(vbKeyQ)) Then
        returnValue = "quit"
        imDone = True
    ElseIf (GetAsyncKeyState(vbKeyUp)) And keyUp Then
        rotateFigure
        keyUp = False
        figureDraw = True
    ElseIf (GetAsyncKeyState(vbKeyDown)) Then
        fallTmr = fallTmr - fallTick
        figureDraw = True
    ElseIf (GetAsyncKeyState(vbKeyLeft)) And keyLeft Then
        keyLeft = False
        fX = fX - 1
        If collWallWell Then fX = fX + 1
        If collAnotherFigure Then fX = fX + 1
        figureDraw = True
    ElseIf (GetAsyncKeyState(vbKeyRight)) And keyRight Then
        keyRight = False
        fX = fX + 1
        If collWallWell Then fX = fX - 1
        If collAnotherFigure Then fX = fX - 1
        figureDraw = True
    ElseIf (GetAsyncKeyState(vbKeySpace)) And keySpace Then
        dropFigure
        keySpace = False
        figureDraw = True
    End If
    
    If (GetAsyncKeyState(vbKeyUp)) = 0 Then
        keyUp = True
    End If
    If (GetAsyncKeyState(vbKeySpace)) = 0 Then
        keySpace = True
    End If
    If (GetAsyncKeyState(vbKeyLeft)) = 0 Then
        keyLeft = True
    End If
    If (GetAsyncKeyState(vbKeyRight)) = 0 Then
        keyRight = True
    End If
End Sub

Private Sub Update()
    If Timer > fallTmr + fallTick Then
        fY = fY - 1
        If collBotWell Then
            fY = fY + 1
            copyFigureToBoard
            nextFigure
        End If
        If collAnotherFigure Then
            fY = fY + 1
            copyFigureToBoard
            nextFigure
        End If
        fallTmr = Timer
        figureDraw = True
    End If
    
    checkLines
End Sub

Private Sub copyFigureToBoard()
    Dim e As Integer, i As Integer
    For i = 0 To UBound(fMatrix, 2)
        For e = 0 To UBound(fMatrix, 1)
            If fMatrix(e, i) = 1 Then
                boardMatrix(e + fX, i + fY) = 1
            End If
        Next e
    Next i
End Sub

Private Sub dropFigure()
    Dim e As Integer, i As Integer
    For i = fY To 0 - UBound(fMatrix, 2) Step -1
        fY = fY - 1
        If collBotWell Then
            fY = fY + 1
            Exit Sub
        End If
        If collAnotherFigure Then
            fY = fY + 1
            Exit Sub
        End If
    Next i
End Sub

Private Sub checkLines()
    Dim e As Integer, i As Integer
    Dim lineIsFull As Boolean
    
    For i = 0 To UBound(boardMatrix, 2)
        lineIsFull = True
        For e = 0 To UBound(boardMatrix, 1)
            If boardMatrix(e, i) = 0 Then
                lineIsFull = False
            End If
        Next e
        If lineIsFull Then
            deleteLine (i)
            i = i - 1
        End If
    Next i
End Sub

Private Sub deleteLine(n As Integer)
    Dim e As Integer, i As Integer
    For e = 0 To UBound(boardMatrix, 1)
        boardMatrix(e, n) = 0
    Next e
    For i = n + 1 To UBound(boardMatrix, 2)
        For e = 0 To UBound(boardMatrix, 1)
            boardMatrix(e, i - 1) = boardMatrix(e, i)
        Next e
    Next i
End Sub

Private Function collAnotherFigure() As Boolean
    Dim e As Integer, i As Integer
    '/ check collision with another figure
    For i = 0 To UBound(fMatrix, 2)
        For e = 0 To UBound(fMatrix, 1)
            If fMatrix(e, i) = 1 Then
                If boardMatrix(fX + e, fY + i) = 1 Then
                    collAnotherFigure = True
                End If
            End If
        Next e
    Next i
End Function

Private Function collBotWell() As Boolean
    Dim e As Integer, i As Integer
    '/ check collision with bottom well
    For i = 0 To UBound(fMatrix, 2)
        For e = 0 To UBound(fMatrix, 1)
            If fMatrix(e, i) = 1 Then
                If fY + i < 0 Then
                    collBotWell = True
                End If
            End If
        Next e
    Next i
End Function

Private Function collWallWell() As Boolean
    Dim e As Integer, i As Integer
    For i = 0 To UBound(fMatrix, 2)
        For e = 0 To UBound(fMatrix, 1)
            If fMatrix(e, i) = 1 Then
                If fX + e < 0 Or fX + e > UBound(boardMatrix, 1) Then
                    collWallWell = True
                End If
            End If
        Next e
    Next i
End Function

Private Sub Draw()
    Application.Optimization = True
    
    '/ draw board
    Dim e As Integer, i As Integer
    Dim s As Shape
    If boardDraw Then
        ActivePage.Layers.Item(3).Shapes.All.Delete
        For i = 0 To (boardHeight - 1)
            For e = 0 To (boardWidth - 1)
                If boardMatrix(e, i) = 1 Then
                    Set s = ActivePage.Layers.Item(3).CreateRectangle(e * cellSize, i * cellSize + cellSize, e * cellSize + cellSize, i * cellSize)
                    s.Outline.SetNoOutline
                    s.Fill.UniformColor.CMYKAssign 100, 0, 0, 0
                End If
            Next e
        Next i
        boardDraw = False
    End If
    '/ draw figure
    If figureDraw Then
        ActivePage.Layers.Item(2).Shapes.All.Delete
        For i = 0 To UBound(fMatrix, 2)
            For e = 0 To UBound(fMatrix, 1)
                If fMatrix(e, i) = 1 Then
                    Set s = ActivePage.Layers.Item(2).CreateRectangle((e + fX) * cellSize, (i + fY) * cellSize + cellSize, (e + fX) * cellSize + cellSize, (i + fY) * cellSize)
                    s.Outline.SetNoOutline
                    s.Fill.UniformColor.CMYKAssign 100, 0, 0, 0
                End If
            Next e
        Next i
        figureDraw = False
    End If
    
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Private Sub Destroy()
    Application.Optimization = True
    
    ActivePage.Layers.Item(2).Shapes.All.Delete
    ActivePage.Layers.Item(3).Shapes.All.Delete
    ActivePage.Layers.Item(4).Shapes.All.Delete
    ActivePage.Layers.Item(5).Shapes.All.Delete
    ActivePage.Layers.Item(6).Shapes.All.Delete
    
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

