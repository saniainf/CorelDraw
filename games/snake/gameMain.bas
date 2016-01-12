Attribute VB_Name = "gameMain"
Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim ImDone As Boolean
Dim directSnake As String
Dim tmr As Double
Dim cellSize As Integer
Dim boardHeight As Integer, boardWidth As Integer
Dim foodMatrix() As Integer
Dim snake() As Integer
Dim debugTxt1 As Shape
Dim Tick As Double
Dim keyReadDone As Boolean
Dim returnValue As Integer

Sub LoadLevel()
    Tick = gameLevel.Tick
    cellSize = gameLevel.cellSize
    boardHeight = gameLevel.boardHeight
    boardWidth = gameLevel.boardWidth
    snake = gameLevel.snake
    foodMatrix = gameLevel.foodMatrix
End Sub

Function GameLoop() As Integer
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
    GameLoop = returnValue
End Function

Private Sub Initialization()
    Dim maxViewArea As Integer
    maxViewArea = 450
    
    ActiveDocument.ActivePage.SetSize (boardWidth) * cellSize, (boardHeight) * cellSize
    ActiveWindow.ActiveView.SetViewArea -(maxViewArea - boardWidth * cellSize) / 2, -(maxViewArea - boardHeight * cellSize) / 2, maxViewArea, maxViewArea
    
    ImDone = False
    directSnake = ""
    keyReadDone = True
    
    drawGameField
End Sub

Private Sub drawGameField()
    Dim s As Shape
    Dim e As Integer, i As Integer
    Application.Optimization = True
    For i = 0 To boardHeight - 1
        For e = 0 To boardWidth - 1
            Set s = ActivePage.Layers.Item(5).CreateRectangle(e * cellSize, i * cellSize, e * cellSize + cellSize, i * cellSize + cellSize)
            s.Fill.ApplyNoFill
            s.Outline.Color.CMYKAssign 0, 0, 0, 40
            s.Outline.Width = 0.1
        Next e
    Next i
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Private Sub UpdateInput()
    If (GetAsyncKeyState(vbKeyQ)) Then
        ImDone = True
    End If
    If (GetAsyncKeyState(vbKeyUp)) And Not directSnake = "down" And Not keyReadDone Then
        directSnake = "up"
        keyReadDone = True
    End If
    If (GetAsyncKeyState(vbKeyDown)) And Not directSnake = "up" And Not keyReadDone Then
        directSnake = "down"
        keyReadDone = True
    End If
    If (GetAsyncKeyState(vbKeyLeft)) And Not directSnake = "right" And Not keyReadDone Then
        directSnake = "left"
        keyReadDone = True
    End If
    If (GetAsyncKeyState(vbKeyRight)) And Not directSnake = "left" And Not keyReadDone Then
        directSnake = "right"
        keyReadDone = True
    End If
End Sub

Private Sub Update()
    Dim a As Integer, b As Integer
    Dim a2 As Integer, b2 As Integer
    Dim e As Integer, i As Integer
    Dim imWin As Boolean
    
    keyReadDone = False
    If directSnake = "" Then Exit Sub
    
    imWin = True
    a = snake(0, 0)
    b = snake(1, 0)
    
    If foodMatrix(a, b) = 1 Then
        ReDim Preserve snake(1, (UBound(snake, 2) + 1))
        foodMatrix(a, b) = 0
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
    
    For e = 1 To UBound(snake, 2)
        a2 = snake(0, e)
        b2 = snake(1, e)
        snake(0, e) = a
        snake(1, e) = b
        a = a2
        b = b2
    Next e
    
    If snake(0, 0) < 0 Or snake(0, 0) > boardWidth - 1 Then
        returnValue = 0
        ImDone = True
    End If
    If snake(1, 0) < 0 Or snake(1, 0) > boardHeight - 1 Then
        returnValue = 0
        ImDone = True
    End If
    For e = 1 To UBound(snake, 2)
        If snake(0, 0) = snake(0, e) And snake(1, 0) = snake(1, e) Then
            returnValue = 0
            ImDone = True
        End If
    Next e
    
    For i = 0 To boardHeight - 1
        For e = 0 To boardWidth - 1
            If foodMatrix(e, i) = 1 Then
                imWin = False
            End If
        Next e
    Next i
    If imWin Then
        returnValue = 1
        ImDone = True
    End If
    
End Sub

Private Sub Draw()
    Application.Optimization = True
    Dim X As Integer, Y As Integer
    Dim e As Integer, i As Integer
    Dim s As Shape
    
    ActivePage.Layers.Item(2).Shapes.All.Delete
    ActivePage.Layers.Item(3).Shapes.All.Delete
    
    X = snake(0, 0) * cellSize
    Y = snake(1, 0) * cellSize
    Set s = ActivePage.Layers.Item(2).CreateEllipse(X, Y + cellSize, X + cellSize, Y)
    s.Outline.SetNoOutline
    s.Fill.UniformColor.CMYKAssign 100, 0, 100, 0
    
    For e = 1 To UBound(snake, 2)
        X = snake(0, e) * cellSize
        Y = snake(1, e) * cellSize
        Set s = ActivePage.Layers.Item(2).CreateEllipse(X, Y + cellSize, X + cellSize, Y)
        s.Outline.SetNoOutline
        s.Fill.UniformColor.CMYKAssign 100, 0, 0, 0
    Next e
    
    For i = 0 To boardHeight - 1
        For e = 0 To boardWidth - 1
            If foodMatrix(e, i) = 1 Then
                Set s = ActivePage.Layers.Item(3).CreateEllipse(e * cellSize, i * cellSize + cellSize, e * cellSize + cellSize, i * cellSize)
                s.Outline.SetNoOutline
                s.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
            End If
        Next e
    Next i
    
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Private Sub Destroy()
    Application.Optimization = True
    
    ActivePage.Layers.Item(2).Shapes.All.Delete
    ActivePage.Layers.Item(3).Shapes.All.Delete
    ActivePage.Layers.Item(5).Shapes.All.Delete
    
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

