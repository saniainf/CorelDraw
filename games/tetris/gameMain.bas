Attribute VB_Name = "gameMain"
Option Explicit

Dim imDone As Boolean
Dim cellSize As Integer
Dim boardHeight As Integer, boardWidth As Integer
Dim screenWidth As Integer, screenHeight As Integer
Dim fallTmr As Double, moveTmr As Double, startMoveTmr As Double
Dim fallTick As Double, moveTick As Double, startMoveTick As Double
Dim keyUp As Boolean, keyDown As Boolean, keyLeft As Boolean, keyRight As Boolean, keySpace As Boolean
Dim returnValue As String
Dim fMatrix() As Integer
Dim boardMatrix() As Integer
Dim fX As Integer, fY As Integer
Dim boardDraw As Boolean, figureDraw As Boolean, interfaceDraw As Boolean
Dim cFigures As New Collection
Dim figO As New Collection, figI As New Collection, figS As New Collection, figZ As New Collection, figL As New Collection, figJ As New Collection, figT As New Collection
Dim currentFigure As Integer, currentShape As Integer, nextFigure As Integer
Public ScorePoint As Integer, currentLvl As Integer, deltaLvlUp As Integer
Dim sScorePoint As Shape, sLevel As Shape

Sub LoadLevel()
    fallTick = 0.9
    startMoveTick = 0.4
    moveTick = 0.06
    cellSize = 15
    boardHeight = 20
    boardWidth = 10
    deltaLvlUp = 10000
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
    interfaceDraw = False
    Set cFigures = New Collection
    Set figO = New Collection
    Set figI = New Collection
    Set figS = New Collection
    Set figZ = New Collection
    Set figL = New Collection
    Set figJ = New Collection
    Set figT = New Collection
    ScorePoint = 0
    currentLvl = 1
    ReDim boardMatrix(boardWidth - 1, boardHeight - 1)
    drawGameField
    ActiveWindow.ActiveView.SetViewArea 0, 0, ActivePage.BoundingBox.Width, ActivePage.BoundingBox.Height
    
    figO.Add strToArr("0,0,0,0,0,1,1,0,0,1,1,0,0,0,0,0") 'yellow
    cFigures.Add figO

    figI.Add strToArr("0,0,0,0,2,2,2,2,0,0,0,0,0,0,0,0") 'cyan
    figI.Add strToArr("0,0,2,0,0,0,2,0,0,0,2,0,0,0,2,0")
    cFigures.Add figI

    figS.Add strToArr("0,0,0,0,0,3,3,0,0,0,3,3,0,0,0,0") 'green
    figS.Add strToArr("0,0,0,0,0,0,0,3,0,0,3,3,0,0,3,0")
    cFigures.Add figS

    figZ.Add strToArr("0,0,0,0,0,0,4,4,0,4,4,0,0,0,0,0")
    figZ.Add strToArr("0,0,0,0,0,0,4,0,0,0,4,4,0,0,0,4")
    cFigures.Add figZ

    figL.Add strToArr("0,0,0,0,0,5,0,0,0,5,5,5,0,0,0,0")
    figL.Add strToArr("0,0,0,0,0,0,5,5,0,0,5,0,0,0,5,0")
    figL.Add strToArr("0,0,0,0,0,0,0,0,0,5,5,5,0,0,0,5")
    figL.Add strToArr("0,0,0,0,0,0,5,0,0,0,5,0,0,5,5,0")
    cFigures.Add figL

    figJ.Add strToArr("0,0,0,0,0,0,0,6,0,6,6,6,0,0,0,0")
    figJ.Add strToArr("0,0,0,0,0,0,6,0,0,0,6,0,0,0,6,6")
    figJ.Add strToArr("0,0,0,0,0,0,0,0,0,6,6,6,0,6,0,0")
    figJ.Add strToArr("0,0,0,0,0,6,6,0,0,0,6,0,0,0,6,0")
    cFigures.Add figJ

    figT.Add strToArr("0,0,0,0,0,0,7,0,0,7,7,7,0,0,0,0") 'purple
    figT.Add strToArr("0,0,0,0,0,0,7,0,0,0,7,7,0,0,7,0")
    figT.Add strToArr("0,0,0,0,0,0,0,0,0,7,7,7,0,0,7,0")
    figT.Add strToArr("0,0,0,0,0,0,7,0,0,7,7,0,0,0,7,0")
    cFigures.Add figT
    
    Set sScorePoint = ActivePage.Layers.Item(6).CreateArtisticText(-100, 250, "0", , , "Arial", 54, cdrTrue, cdrFalse, cdrNoFontLine, cdrLeftAlignment)
    Set sLevel = ActivePage.Layers.Item(6).CreateArtisticText(-100, 200, "Level: " & currentLvl, , , "Arial", 54, cdrTrue, cdrFalse, cdrNoFontLine, cdrLeftAlignment)
            
    Randomize
    nextFigure = Int((cFigures.Count * Rnd) + 1)
            
    placeNextFigure
End Sub

Private Sub placeNextFigure()
    Dim e As Integer, i As Integer
    
    Randomize
    currentFigure = nextFigure
    nextFigure = Int((cFigures.Count * Rnd) + 1)
    currentShape = 1
    fMatrix = cFigures.Item(currentFigure).Item(currentShape)
    interfaceDraw = True
    
    fallTmr = Timer
    fX = 3
    fY = 0
    For i = UBound(fMatrix, 2) To 0 Step -1
        For e = 0 To UBound(fMatrix, 1)
            If Not fMatrix(e, i) = 0 Then
                fY = UBound(boardMatrix, 2) - i
                Exit For
            End If
        Next e
        If fY > 0 Then
            Exit For
        End If
    Next i
    
    If checkCollision Then
        returnValue = "gameover"
        imDone = True
        Exit Sub
    End If
    figureDraw = True
End Sub

Private Sub rotateFigure()
    Dim prevShape As Integer
    
    prevShape = currentShape
    currentShape = currentShape + 1
    If currentShape > cFigures.Item(currentFigure).Count Then
        currentShape = 1
    End If
    fMatrix = cFigures.Item(currentFigure).Item(currentShape)
    
    If checkCollision Then
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
            s.Outline.Width = 0.1
        Next e
    Next i
    ActiveDocument.ClearSelection
End Sub

Private Sub UpdateInput()
    If (GetAsyncKeyState(vbKeyQ)) Then
        returnValue = "quit"
        imDone = True
    ElseIf (GetAsyncKeyState(vbKeyUp)) And keyUp Then 'rotate
        rotateFigure
        keyUp = False
    ElseIf (GetAsyncKeyState(vbKeyDown)) Then 'fast fall
        If keyDown Then
            keyDown = False
            fY = fY - 1
            If checkCollision Then
                fY = fY + 1
            Else
                fallTmr = Timer
            End If
            boardDraw = True
            figureDraw = True
            startMoveTmr = Timer
        End If
        If Timer > startMoveTmr + startMoveTick Then
            If Timer > moveTmr + moveTick Then
                fY = fY - 1
                    If checkCollision Then
                        fY = fY + 1
                    Else
                        fallTmr = Timer
                    End If
                moveTmr = Timer
                figureDraw = True
            End If
        End If
    ElseIf (GetAsyncKeyState(vbKeyLeft)) Then   'left
        If keyLeft Then
            keyLeft = False
            fX = fX - 1
            If checkCollision Then fX = fX + 1
            figureDraw = True
            startMoveTmr = Timer
        End If
        If Timer > startMoveTmr + startMoveTick Then
            If Timer > moveTmr + moveTick Then
                fX = fX - 1
                If checkCollision Then fX = fX + 1
                moveTmr = Timer
                figureDraw = True
            End If
        End If
    ElseIf (GetAsyncKeyState(vbKeyRight)) Then  'right
        If keyRight Then
            keyRight = False
            fX = fX + 1
            If checkCollision Then fX = fX - 1
            figureDraw = True
            startMoveTmr = Timer
        End If
        If Timer > startMoveTmr + startMoveTick Then
            If Timer > moveTmr + moveTick Then
                fX = fX + 1
                If checkCollision Then fX = fX - 1
                moveTmr = Timer
                figureDraw = True
            End If
        End If
    ElseIf (GetAsyncKeyState(vbKeySpace)) And keySpace Then 'drop
        dropFigure
        keySpace = False
    End If
    
    If (GetAsyncKeyState(vbKeyUp)) = 0 Then
        keyUp = True
    End If
    If (GetAsyncKeyState(vbKeyDown)) = 0 Then
        keyDown = True
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
        If checkCollision Then
            fY = fY + 1
            copyFigureToBoard
            placeNextFigure
        End If
        fallTmr = Timer
        figureDraw = True
    End If
    If checkLines Then boardDraw = True
    If ScorePoint >= currentLvl * deltaLvlUp Then
        fallTick = fallTick - 0.02
        currentLvl = currentLvl + 1
        interfaceDraw = True
    End If
End Sub

Private Sub copyFigureToBoard()
    Dim e As Integer, i As Integer
    For i = 0 To UBound(fMatrix, 2)
        For e = 0 To UBound(fMatrix, 1)
            If Not fMatrix(e, i) = 0 Then
                boardMatrix(e + fX, i + fY) = fMatrix(e, i)
            End If
        Next e
    Next i
    boardDraw = True
End Sub

Private Sub dropFigure()
    Dim e As Integer, i As Integer
    For i = fY To 0 - UBound(fMatrix, 2) Step -1
        fY = fY - 1
        If checkCollision Then
            fY = fY + 1
            figureDraw = True
            fallTmr = Timer
            Exit Sub
        End If
    Next i
End Sub

Private Function checkLines() As Boolean
    Dim lineCount As Integer
    Dim e As Integer, i As Integer
    Dim lineIsFull As Boolean
    lineCount = 0
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
            checkLines = True
            lineCount = lineCount + 1
        End If
    Next i
    If lineCount > 0 Then
        Select Case lineCount
            Case 1
                ScorePoint = ScorePoint + 100
            Case 2
                ScorePoint = ScorePoint + 300
            Case 3
                ScorePoint = ScorePoint + 700
            Case 4
                ScorePoint = ScorePoint + 1500
            Case Else
        End Select
        interfaceDraw = True
    End If
End Function

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

Private Function checkCollision() As Boolean
    Dim e As Integer, i As Integer
    checkCollision = False
    For i = 0 To UBound(fMatrix, 2)
        For e = 0 To UBound(fMatrix, 1)
            If Not fMatrix(e, i) = 0 Then
                '/ check collision with wall of the well
                If fX + e < 0 Or fX + e > UBound(boardMatrix, 1) Then
                    checkCollision = True
                '/ check collision with bottom of the well
                ElseIf fY + i < 0 Then
                    checkCollision = True
                '/ check collision with top of the well
                ElseIf fY + i > UBound(boardMatrix, 2) Then
                    checkCollision = True
                '/ check collision with dropped figures
                ElseIf Not boardMatrix(fX + e, fY + i) = 0 Then
                    checkCollision = True
                End If
            End If
        Next e
    Next i
End Function

Private Sub Draw()
    Dim e As Integer, i As Integer
    Dim s As Shape
    Application.Optimization = True
    '/ draw interface
    If interfaceDraw Then
        sScorePoint.Text.Story = ScorePoint
        sLevel.Text.Story = "Level: " & currentLvl
        ActivePage.Layers.Item(4).Shapes.All.Delete
        For i = 0 To UBound(cFigures.Item(nextFigure).Item(1), 2)
            For e = 0 To UBound(cFigures.Item(nextFigure).Item(1), 1)
                If Not cFigures.Item(nextFigure).Item(1)(e, i) = 0 Then
                    Set s = ActivePage.Layers.Item(4).CreateRectangle(180 + e * cellSize, 180 + i * cellSize + cellSize, 180 + e * cellSize + cellSize, 180 + i * cellSize, 20, 20, 20, 20)
                    s.Outline.Color.CMYKAssign 0, 0, 0, 100
                    s.Outline.Width = 0.4
                    Select Case cFigures.Item(nextFigure).Item(1)(e, i)
                        Case 1
                            s.Fill.UniformColor.CMYKAssign 0, 0, 100, 0
                        Case 2
                            s.Fill.UniformColor.CMYKAssign 100, 0, 0, 0
                        Case 3
                            s.Fill.UniformColor.CMYKAssign 100, 0, 100, 0
                        Case 4
                            s.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
                        Case 5
                            s.Fill.UniformColor.CMYKAssign 0, 60, 100, 0
                        Case 6
                            s.Fill.UniformColor.CMYKAssign 100, 100, 0, 0
                        Case 7
                            s.Fill.UniformColor.CMYKAssign 40, 100, 0, 0
                    End Select
                End If
            Next e
        Next i
        interfaceDraw = False
    End If
    '/ draw board
    If boardDraw Then
        ActivePage.Layers.Item(3).Shapes.All.Delete
        For i = 0 To (boardHeight - 1)
            For e = 0 To (boardWidth - 1)
                If Not boardMatrix(e, i) = 0 Then
                    Set s = ActivePage.Layers.Item(3).CreateRectangle(e * cellSize, i * cellSize + cellSize, e * cellSize + cellSize, i * cellSize, 20, 20, 20, 20)
                    s.Outline.Color.CMYKAssign 0, 0, 0, 100
                    s.Outline.Width = 0.4
                    Select Case boardMatrix(e, i)
                        Case 1
                            s.Fill.UniformColor.CMYKAssign 0, 0, 100, 0
                        Case 2
                            s.Fill.UniformColor.CMYKAssign 100, 0, 0, 0
                        Case 3
                            s.Fill.UniformColor.CMYKAssign 100, 0, 100, 0
                        Case 4
                            s.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
                        Case 5
                            s.Fill.UniformColor.CMYKAssign 0, 60, 100, 0
                        Case 6
                            s.Fill.UniformColor.CMYKAssign 100, 100, 0, 0
                        Case 7
                            s.Fill.UniformColor.CMYKAssign 40, 100, 0, 0
                    End Select
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
                If Not fMatrix(e, i) = 0 Then
                    Set s = ActivePage.Layers.Item(2).CreateRectangle((e + fX) * cellSize, (i + fY) * cellSize + cellSize, (e + fX) * cellSize + cellSize, (i + fY) * cellSize, 20, 20, 20, 20)
                    s.Outline.Color.CMYKAssign 0, 0, 0, 100
                    s.Outline.Width = 0.4
                    Select Case fMatrix(e, i)
                        Case 1
                            s.Fill.UniformColor.CMYKAssign 0, 0, 100, 0
                        Case 2
                            s.Fill.UniformColor.CMYKAssign 100, 0, 0, 0
                        Case 3
                            s.Fill.UniformColor.CMYKAssign 100, 0, 100, 0
                        Case 4
                            s.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
                        Case 5
                            s.Fill.UniformColor.CMYKAssign 0, 60, 100, 0
                        Case 6
                            s.Fill.UniformColor.CMYKAssign 100, 100, 0, 0
                        Case 7
                            s.Fill.UniformColor.CMYKAssign 40, 100, 0, 0
                    End Select
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

