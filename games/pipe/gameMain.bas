Attribute VB_Name = "gameMain"
Option Explicit

Dim imDone As Boolean
Dim cellSize As Integer
Dim boardHeight As Integer, boardWidth As Integer
Dim screenWidth As Integer, screenHeight As Integer
Dim keyReadDone As Boolean
Dim returnValue As String
Dim drawBoard As Boolean
Dim gcPipe As New Collection, gcWaterPipe As New Collection
Dim pipeTypes() As Variant
Dim gameBoard() As Integer
Dim waterPipes() As Integer

Sub LoadLevel()
    cellSize = 10
    boardHeight = 10
    boardWidth = 8
End Sub

Function GameLoop() As String
    LoadResource
    Initialization
    Do Until imDone
        DoEvents
        UpdateInput
        Draw
        Update
    Loop
    Destroy
    GameLoop = returnValue
End Function

Private Sub LoadResource()
    Dim s As Shape
    Dim i As Integer
    Set gcPipe = New Collection
    Set gcWaterPipe = New Collection
    For i = 1 To 6
        gcPipe.Add ActiveDocument.Pages.Item(2).Shapes.Item(i)
    Next i
    For i = 7 To 12
        gcWaterPipe.Add ActiveDocument.Pages.Item(2).Shapes.Item(i)
    Next i
    
End Sub

Private Sub Initialization()
    imDone = False
    keyReadDone = True
    drawBoard = False
    
    pipeTypes = Array("Left,Right", "Top,Bottom", "Top,Right", "Bottom,Right", "Top,Left", "Bottom,Left")
    ReDim gameBoard(boardWidth, boardHeight)
    ReDim waterPipes(boardWidth, boardHeight)
    
    drawGameField
    fillGameBoard
    clearWaterPath
    buildWaterPath
    drawBoard = True
End Sub

Private Sub waterPath(x As Integer, y As Integer, inputDirection As String)
    Dim s As Variant
    If x <= boardWidth And x >= 0 And y <= boardHeight And y >= 0 Then
        If hasInPipe(inputDirection, gameBoard(x, y)) = True And waterPipes(x, y) = 0 Then
            waterPipes(x, y) = 1
            For Each s In getOutPipe(gameBoard(x, y), inputDirection)
                Select Case s
                    Case "Left"
                        waterPath x - 1, y, "Right"
                    Case "Right"
                        waterPath x + 1, y, "Left"
                    Case "Top"
                        waterPath x, y + 1, "Bottom"
                    Case "Bottom"
                        waterPath x, y - 1, "Top"
                End Select
            Next s
        End If
    End If
End Sub

Private Sub buildWaterPath()
    Dim i As Integer, e As Integer
    e = 0
    For i = 0 To boardHeight
        waterPath e, i, "Left"
    Next i
End Sub

Private Function hasInPipe(direction As String, typePipe As Integer) As Boolean
    If Not InStr(pipeTypes(typePipe), direction) = 0 Then hasInPipe = True
End Function

Private Function getOutPipe(pipeType As Integer, inputPipe As String) As String()
    Dim s As Variant
    Dim c As New Collection
    Dim i As Integer
    Dim a() As String
    For Each s In Split(pipeTypes(pipeType), ",")
        If Not s = inputPipe Then
            c.Add s
        End If
    Next s
    ReDim a(c.Count)
    For i = 0 To (UBound(a, 1) - 1)
        a(i) = c.Item(i + 1)
    Next i
    getOutPipe = a
End Function

Private Sub clearWaterPath()
    ReDim waterPipes(boardWidth, boardHeight)
End Sub

Private Sub fillGameBoard()
    Dim i As Integer, e As Integer
    Randomize
    For i = 0 To boardHeight
        For e = 0 To boardWidth
            gameBoard(e, i) = Int(Rnd * UBound(pipeTypes, 1))
        Next e
    Next i
End Sub

Private Sub rotatePipePiece(x As Integer, y As Integer)
    Select Case gameBoard(x, y)
        Case 0
            gameBoard(x, y) = 1
        Case 1
            gameBoard(x, y) = 0
        Case 2
            gameBoard(x, y) = 3
        Case 3
            gameBoard(x, y) = 5
        Case 4
            gameBoard(x, y) = 2
        Case 5
            gameBoard(x, y) = 4
        Case 6
    End Select
End Sub

Private Sub drawGameField()
    Dim s As Shape
    Dim e As Integer, i As Integer
    Application.Optimization = True
    For i = 0 To boardHeight - 1
        For e = 0 To boardWidth - 1
            Set s = ActivePage.Layers.Item(4).CreateRectangle(e * cellSize, i * cellSize, e * cellSize + cellSize, i * cellSize + cellSize)
            s.Fill.ApplyNoFill
            s.Outline.Color.CMYKAssign 0, 0, 0, 10
            s.Outline.width = 0.1
        Next e
    Next i
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Private Sub UpdateInput()
    If (GetAsyncKeyState(vbKeyQ)) Then
        returnValue = "quit"
        imDone = True
        keyReadDone = True
    ElseIf (GetAsyncKeyState(vbKeyUp)) And Not keyReadDone Then

        keyReadDone = True
    ElseIf (GetAsyncKeyState(vbKeyDown)) And Not keyReadDone Then

        keyReadDone = True
    ElseIf (GetAsyncKeyState(vbKeyLeft)) And Not keyReadDone Then
        
        keyReadDone = True
    ElseIf (GetAsyncKeyState(vbKeyRight)) And Not keyReadDone Then
        
        keyReadDone = True
    End If
    mouseUpadate
End Sub

Private Sub mouseUpadate()
    Dim x As Double, y As Double
    Dim shift As Long
    ActiveDocument.GetUserClick x, y, shift, 10, False, cdrCursorPick
    rotatePipePiece x \ cellSize, y \ cellSize
    clearWaterPath
    buildWaterPath
    drawBoard = True
End Sub

Private Sub Update()
    keyReadDone = False
End Sub

Private Sub Draw()
    If drawBoard Then
        Dim i As Integer, e As Integer
        Dim s As Shape
        Application.Optimization = True
        ActivePage.Layers.Item(2).Shapes.All.Delete
        ActivePage.Layers.Item(3).Shapes.All.Delete
        
        For i = 0 To boardHeight - 1
            For e = 0 To boardWidth - 1
                If waterPipes(e, i) = 0 Then
                    Set s = gcPipe.Item(gameBoard(e, i) + 1).Duplicate
                    s.MoveToLayer ActivePage.Layers.Item(2)
                    s.SetPosition e * cellSize, i * cellSize
                Else
                    Set s = gcWaterPipe.Item(gameBoard(e, i) + 1).Duplicate
                    s.MoveToLayer ActivePage.Layers.Item(2)
                    s.SetPosition e * cellSize, i * cellSize
                End If
            Next e
        Next i
        
        ActiveDocument.ClearSelection
        Application.Optimization = False
        ActiveWindow.Refresh
        Application.Refresh
        drawBoard = False
    End If
End Sub

Private Sub Destroy()
    Application.Optimization = True
    
    ActivePage.Layers.Item(2).Shapes.All.Delete
    ActivePage.Layers.Item(3).Shapes.All.Delete
    ActivePage.Layers.Item(4).Shapes.All.Delete
    
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub
