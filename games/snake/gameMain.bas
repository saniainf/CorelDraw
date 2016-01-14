Attribute VB_Name = "gameMain"
Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public scorePoint As Integer
Dim ImDone As Boolean
Dim directSnake As String
Dim tmr As Double
Dim cellSize As Integer
Dim boardHeight As Integer, boardWidth As Integer
Dim offsetLeft As Double, offsetBottom As Double
Dim screenWidth As Integer, screenHeight As Integer
Dim foodMatrix() As Integer
Dim wallmatrix() As Integer
Dim snake() As Integer
Dim Tick As Double
Dim keyReadDone As Boolean
Dim returnValue As String
Dim SScorePoint As Shape
Dim gc As New Collection

Sub LoadLevel()
    Tick = gameLevel.Tick
    cellSize = gameLevel.cellSize
    boardHeight = gameLevel.boardHeight
    boardWidth = gameLevel.boardWidth
    snake = gameLevel.snake
    foodMatrix = gameLevel.foodMatrix
    wallmatrix = gameLevel.wallmatrix
    scorePoint = 0
End Sub

Function GameLoop() As String
    Initialization
    LoadResource
    tmr = Timer
    Do Until ImDone
        DoEvents
        UpdateInput
        If Timer > tmr + Tick And Not ImDone Then
            Update
            Draw
            tmr = Timer
        End If
    Loop
    Destroy
    GameLoop = returnValue
End Function

Private Sub LoadResource()
    
End Sub

Private Sub Initialization()
    Dim maxViewArea As Integer
    maxViewArea = 450
    screenWidth = 800
    screenHeight = 450
    
    ActiveDocument.ActivePage.SetSize screenWidth, screenHeight
    ActiveWindow.ActiveView.SetViewArea 0, 0, screenWidth, screenHeight
    offsetLeft = (screenWidth - boardWidth * cellSize) / 2
    offsetBottom = (screenHeight - boardHeight * cellSize) / 2
    
    ImDone = False
    directSnake = ""
    keyReadDone = True
    
    drawGameField
    drawWall
    drawInterface
End Sub

Private Sub drawInterface()
    Set SScorePoint = ActivePage.Layers.Item(6).CreateArtisticText(40, 350, "0", , , "Arial", 54, cdrTrue, cdrFalse, cdrNoFontLine, cdrLeftAlignment)
End Sub

Private Sub drawWall()
    Dim s As Shape
    Dim e As Integer, i As Integer
    Application.Optimization = True
    For i = 0 To boardHeight - 1
        For e = 0 To boardWidth - 1
            If wallmatrix(e, i) = 1 Then
                Set s = ActivePage.Layers.Item(4).CreateRectangle(e * cellSize + offsetLeft, i * cellSize + offsetBottom, e * cellSize + cellSize + offsetLeft, i * cellSize + cellSize + offsetBottom)
                s.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
                s.Outline.SetNoOutline
            End If
        Next e
    Next i
    ActiveDocument.ClearSelection
End Sub

Private Sub drawGameField()
    Dim s As Shape
    Dim e As Integer, i As Integer
    Application.Optimization = True
    For i = 0 To boardHeight - 1
        For e = 0 To boardWidth - 1
            Set s = ActivePage.Layers.Item(5).CreateRectangle(e * cellSize + offsetLeft, i * cellSize + offsetBottom, e * cellSize + cellSize + offsetLeft, i * cellSize + cellSize + offsetBottom)
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
        ImDone = True
        keyReadDone = True
    ElseIf (GetAsyncKeyState(vbKeyUp)) And Not directSnake = "down" And Not keyReadDone Then
        directSnake = "up"
        keyReadDone = True
    ElseIf (GetAsyncKeyState(vbKeyDown)) And Not directSnake = "up" And Not keyReadDone Then
        directSnake = "down"
        keyReadDone = True
    ElseIf (GetAsyncKeyState(vbKeyLeft)) And Not directSnake = "right" And Not keyReadDone Then
        directSnake = "left"
        keyReadDone = True
    ElseIf (GetAsyncKeyState(vbKeyRight)) And Not directSnake = "left" And Not keyReadDone Then
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
    
    '/ collision food
    If foodMatrix(a, b) = 1 Then
        ReDim Preserve snake(1, (UBound(snake, 2) + 1))
        snake(0, UBound(snake, 2)) = a
        snake(1, UBound(snake, 2)) = b
        foodMatrix(a, b) = 0
        scorePoint = scorePoint + 50
    End If
    scorePoint = scorePoint + 1
    
    '/ move head
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
    '/ move body
    For e = 1 To UBound(snake, 2)
        a2 = snake(0, e)
        b2 = snake(1, e)
        snake(0, e) = a
        snake(1, e) = b
        a = a2
        b = b2
    Next e
    
    '/ out of range
    If snake(0, 0) < 0 Or snake(0, 0) > boardWidth - 1 Then
        returnValue = "loselevel"
        ImDone = True
        Exit Sub
    End If
    If snake(1, 0) < 0 Or snake(1, 0) > boardHeight - 1 Then
        returnValue = "loselevel"
        ImDone = True
        Exit Sub
    End If
    '/ collision wall
    If wallmatrix(snake(0, 0), snake(1, 0)) = 1 Then
        returnValue = "loselevel"
        ImDone = True
        Exit Sub
    End If
    '/ collision his body
    For e = 1 To UBound(snake, 2)
        If snake(0, 0) = snake(0, e) And snake(1, 0) = snake(1, e) Then
            returnValue = "loselevel"
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
        returnValue = "endlevel"
        ImDone = True
    End If
    
End Sub

Private Sub Draw()
    Application.Optimization = True
    Dim X As Integer, Y As Integer
    Dim e As Integer, i As Integer
    Dim s As Shape
    Dim typeBodyCell As String
    
    ActivePage.Layers.Item(2).Shapes.All.Delete
    ActivePage.Layers.Item(3).Shapes.All.Delete
    SScorePoint.Text.Story = " "
    '/ draw snake head
    X = snake(0, 0) * cellSize
    Y = snake(1, 0) * cellSize
    Set s = ActivePage.Layers.Item(2).CreateEllipse(X + offsetLeft, Y + cellSize + offsetBottom, X + cellSize + offsetLeft, Y + offsetBottom)
    s.Outline.SetNoOutline
    s.Fill.UniformColor.CMYKAssign 100, 0, 100, 0
    '/ draw snake body
    For e = 1 To UBound(snake, 2)
        If e = UBound(snake, 2) Then
            typeBodyCell = "lr"
        Else
            typeBodyCell = findTypeBodyCell(snake(0, e - 1), snake(1, e - 1), snake(0, e), snake(1, e), snake(0, e + 1), snake(1, e + 1))
        End If
        X = snake(0, e) * cellSize
        Y = snake(1, e) * cellSize
        Select Case typeBodyCell
            Case "tr"
                Set s = ActivePage.Layers.Item(2).CreateEllipse(X + offsetLeft, Y + cellSize + offsetBottom, X + cellSize + offsetLeft, Y + offsetBottom)
                s.Outline.SetNoOutline
                s.Fill.UniformColor.CMYKAssign 50, 0, 0, 0
            Case "br"
                Set s = ActivePage.Layers.Item(2).CreateEllipse(X + offsetLeft, Y + cellSize + offsetBottom, X + cellSize + offsetLeft, Y + offsetBottom)
                s.Outline.SetNoOutline
                s.Fill.UniformColor.CMYKAssign 0, 50, 0, 0
            Case "tl"
                Set s = ActivePage.Layers.Item(2).CreateEllipse(X + offsetLeft, Y + cellSize + offsetBottom, X + cellSize + offsetLeft, Y + offsetBottom)
                s.Outline.SetNoOutline
                s.Fill.UniformColor.CMYKAssign 0, 0, 50, 0
            Case "bl"
                Set s = ActivePage.Layers.Item(2).CreateEllipse(X + offsetLeft, Y + cellSize + offsetBottom, X + cellSize + offsetLeft, Y + offsetBottom)
                s.Outline.SetNoOutline
                s.Fill.UniformColor.CMYKAssign 0, 0, 0, 50
            Case Else
                Set s = ActivePage.Layers.Item(2).CreateEllipse(X + offsetLeft, Y + cellSize + offsetBottom, X + cellSize + offsetLeft, Y + offsetBottom)
                s.Outline.SetNoOutline
                s.Fill.UniformColor.CMYKAssign 100, 0, 0, 0
        End Select
    Next e
    '/ draw food
    For i = 0 To boardHeight - 1
        For e = 0 To boardWidth - 1
            If foodMatrix(e, i) = 1 Then
                Set s = ActivePage.Layers.Item(3).CreateEllipse(e * cellSize + offsetLeft, i * cellSize + cellSize + offsetBottom, e * cellSize + cellSize + offsetLeft, i * cellSize + offsetBottom)
                s.Outline.SetNoOutline
                s.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
            End If
        Next e
    Next i
    '/ draw interface
    SScorePoint.Text.Story = scorePoint
    
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Private Function findTypeBodyCell(pX As Integer, pY As Integer, X As Integer, Y As Integer, nX As Integer, nY As Integer) As String
    Dim a As String
    Dim b As String
    
    If X = pX Then
        If pY = Y + 1 Then a = "top"
        If pY = Y - 1 Then a = "bottom"
    End If
    If Y = pY Then
        If pX = X + 1 Then a = "right"
        If pX = X - 1 Then a = "left"
    End If
    
    If X = nX Then
        If nY = Y + 1 Then b = "top"
        If nY = Y - 1 Then b = "bottom"
    End If
    If Y = nY Then
        If nX = X + 1 Then b = "right"
        If nX = X - 1 Then b = "left"
    End If
    
    If (a = "top" And b = "right") Or (a = "right" And b = "top") Then
        findTypeBodyCell = "tr"
    End If
    If (a = "bottom" And b = "right") Or (a = "right" And b = "bottom") Then
        findTypeBodyCell = "br"
    End If
    If (a = "top" And b = "left") Or (a = "left" And b = "top") Then
        findTypeBodyCell = "tl"
    End If
    If (a = "bottom" And b = "left") Or (a = "left" And b = "bottom") Then
        findTypeBodyCell = "bl"
    End If
End Function

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

