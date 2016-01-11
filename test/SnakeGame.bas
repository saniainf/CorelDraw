Attribute VB_Name = "SnakeGame"
Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim ImDone As Boolean
Dim EndGame As Boolean
Dim directSnake As String
Dim tmr As Double
Dim cellSize As Integer
Dim boardHeight As Integer, boardWidth As Integer
Dim matrix() As Integer
Dim snake() As Integer
Dim debugTxt1 As Shape
Dim Tick As Double
Dim foodCount As Integer

Sub SnakeGame()
    Application.ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrTopLeft
    EndGame = False
    

End Sub

Sub SnakeGameLoop()
    Initialization
    tmr = Timer
    Do Until ImDone
        DoEvents
        UpdateInput
        If Timer > tmr + Tick Then
            Update
            Draw
            tmr = Timer
        End If
    Loop
    Destroy
End Sub

Sub Initialization()
    Dim i As Integer, e As Integer
    Dim startBodySize As Integer
    Dim startRow As Integer, startColumn As Integer
    Dim s As Shape
    Dim X As Integer, Y As Integer
    
    cellSize = 15
    startBodySize = 5 - 1
    boardHeight = 20 - 1
    boardWidth = 20 - 1
    startRow = 20 - 1
    startColumn = 2 - 1
    foodCount = 5 - 1
    
    Tick = 0.3
    
    ReDim matrix(boardWidth, boardHeight)
    ReDim snake(1, startBodySize)
    
    ActiveDocument.ActivePage.SetSize (boardWidth + 1) * cellSize, (boardWidth + 1) * cellSize
    
    ImDone = False
    directSnake = "right"
    
    drawGameField
    
    For e = startBodySize To 0 Step -1
        snake(0, e) = Abs(e - startBodySize) + startColumn
        snake(1, e) = startRow
    Next e
    
    Randomize
    For e = 0 To foodCount
        X = Int(20 * Rnd)
        Y = Int(20 * Rnd)
        matrix(X, Y) = 1
    Next e
    
End Sub

Sub drawGameField()
    Dim s As Shape
    Dim e As Integer, i As Integer
    Application.Optimization = True
    For i = 0 To boardHeight
        For e = 0 To boardWidth
            Set s = ActivePage.Layers.Item(5).CreateRectangle(e * cellSize, i * cellSize, e * cellSize + cellSize, i * cellSize + cellSize)
            s.Fill.ApplyNoFill
            s.Outline.Color.CMYKAssign 0, 0, 0, 40
            s.Outline.width = 0.1
        Next e
    Next i
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Sub UpdateInput()
    If (GetAsyncKeyState(vbKeyQ)) Then
        ImDone = True
        EndGame = True
    End If
    If (GetAsyncKeyState(vbKeyUp)) And Not directSnake = "down" Then
        directSnake = "up"
    End If
    If (GetAsyncKeyState(vbKeyDown)) And Not directSnake = "up" Then
        directSnake = "down"
    End If
    If (GetAsyncKeyState(vbKeyLeft)) And Not directSnake = "right" Then
        directSnake = "left"
    End If
    If (GetAsyncKeyState(vbKeyRight)) And Not directSnake = "left" Then
        directSnake = "right"
    End If
End Sub

Sub Update()
    Dim a As Integer, b As Integer
    Dim a2 As Integer, b2 As Integer
    Dim e As Integer, i As Integer
    Dim imWin As Boolean
    
    imWin = True
    a = snake(0, 0)
    b = snake(1, 0)
    
    If matrix(a, b) = 1 Then
        ReDim Preserve snake(1, (UBound(snake, 2) + 1))
        matrix(a, b) = 0
    End If
    
    Select Case directSnake
        Case "right"
        snake(0, 0) = snake(0, 0) + 1
        Case "left"
        snake(0, 0) = snake(0, 0) - 1
        Case "up"
        snake(1, 0) = snake(1, 0) + 1
        Case "down"
        snake(1, 0) = snake(1, 0) - 1
    End Select
    
    If snake(0, 0) < 0 Or snake(0, 0) > boardWidth Then
        gameOver
    End If
    If snake(1, 0) < 0 Or snake(1, 0) > boardHeight Then
        gameOver
    End If
    
    For e = 1 To UBound(snake, 2)
        a2 = snake(0, e)
        b2 = snake(1, e)
        snake(0, e) = a
        snake(1, e) = b
        a = a2
        b = b2
    Next e
    
    For e = 1 To UBound(snake, 2)
        If snake(0, 0) = snake(0, e) And snake(1, 0) = snake(1, e) Then
            gameOver
        End If
    Next e
    
    For i = 0 To boardWidth
        For e = 0 To boardHeight
            If matrix(e, i) = 1 Then
                imWin = False
            End If
        Next e
    Next i
    
    If imWin Then gameWin
    
End Sub

Sub Draw()
    Application.Optimization = True
    Dim X As Integer, Y As Integer
    Dim e As Integer, i As Integer
    Dim s As Shape
    
    ActivePage.Layers.Item(2).Shapes.All.Delete
    ActivePage.Layers.Item(3).Shapes.All.Delete
    
    X = snake(0, 0) * cellSize
    Y = snake(1, 0) * cellSize
    Set s = ActivePage.Layers.Item(2).CreateRectangle(X, Y + cellSize, X + cellSize, Y)
    s.Fill.UniformColor.CMYKAssign 100, 0, 100, 0
    
    For e = 1 To UBound(snake, 2)
        X = snake(0, e) * cellSize
        Y = snake(1, e) * cellSize
        Set s = ActivePage.Layers.Item(2).CreateRectangle(X, Y + cellSize, X + cellSize, Y)
        s.Fill.UniformColor.CMYKAssign 100, 0, 0, 0
    Next e
    
    For i = 0 To boardWidth
        For e = 0 To boardHeight
            If matrix(e, i) = 1 Then
                Set s = ActivePage.Layers.Item(3).CreateRectangle(e * cellSize, i * cellSize + cellSize, e * cellSize + cellSize, i * cellSize)
                s.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
            End If
        Next e
    Next i
    
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Sub Destroy()
    Application.Optimization = True
    
    ActivePage.Layers.Item(2).Shapes.All.Delete
    ActivePage.Layers.Item(3).Shapes.All.Delete
    ActivePage.Layers.Item(5).Shapes.All.Delete
    
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Sub gameOver()
    MsgBox "Game Over"
    ImDone = True
End Sub

Sub gameWin()
    MsgBox "You Win"
    ImDone = True
End Sub

