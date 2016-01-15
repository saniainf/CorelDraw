Attribute VB_Name = "mMain"
Option Explicit
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim imDone As Boolean

Sub SnakeGame()
    Application.ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrTopLeft
    Dim i As Integer
    Dim returnValue As String
    Dim scorePoint As Integer
    scorePoint = 0
    
    newGame
    
    For i = 1 To gameLevel.levelCount
        ActiveDocument.Pages.Item(2).Activate
        gameLevel.levelChange (i)
        gameMain.LoadLevel
        gameMain.scorePoint = scorePoint
        returnValue = gameMain.GameLoop
        scorePoint = gameMain.scorePoint
        Select Case returnValue
            Case "loselevel"
                gameOver
                Exit Sub
            Case "endlevel"
                If Not i = gameLevel.levelCount Then
                    nextLevel (i + 1)
                Else
                    gameWin
                End If
            Case "quit"
                Exit Sub
        End Select
    Next i
End Sub

Private Sub newGame()
    imDone = False
    
    ActiveDocument.Pages.Item(1).Activate
    Do Until imDone
        DoEvents
        UpdateInput
    Loop
    nextLevel (1)
End Sub

Private Sub gameOver()
    Dim s As Shape
    ActiveDocument.Pages.Item(6).Activate
    Set s = ActivePage.Layers.Item(2).CreateArtisticText(400, 170, scorePoint, , , "Arial", 84, cdrTrue, cdrFalse, cdrNoFontLine, cdrCenterAlignment)
    s.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
    ActiveDocument.ClearSelection
End Sub

Private Sub nextLevel(i As Integer)
    Dim s As Shape
    Dim tmr As Double
    
    imDone = False
    
    Application.Optimization = True
    ActiveDocument.Pages.Item(4).Activate
    Set s = ActivePage.Layers.Item(2).CreateArtisticText(400, 220, "Level " & i, , , "Arial", 84, cdrTrue, cdrFalse, cdrNoFontLine, cdrCenterAlignment)
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
    
    tmr = Timer
    Do Until imDone
        DoEvents
        If Timer > tmr + 2 Then
            imDone = True
        End If
    Loop
    
    Application.Optimization = True
    s.Delete
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Private Sub gameWin()
    Dim s As Shape
    ActiveDocument.Pages.Item(3).Activate
    Set s = ActivePage.Layers.Item(2).CreateArtisticText(400, 170, scorePoint, , , "Arial", 84, cdrTrue, cdrFalse, cdrNoFontLine, cdrCenterAlignment)
    s.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
    ActiveDocument.ClearSelection
End Sub

Private Sub UpdateInput()
    If (GetAsyncKeyState(vbKeySpace)) Then
          imDone = True
    End If
End Sub
