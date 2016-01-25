Attribute VB_Name = "mMain"
Option Explicit
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim imDone As Boolean

Sub TetrisGame()
    Application.ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft
    Dim returnValue As String
    
'    newGame
    ActiveDocument.Pages.Item(2).Activate
    ActiveDocument.MasterPage.DesktopLayer.Visible = True
    gameMain.LoadLevel
    returnValue = gameMain.GameLoop
    Select Case returnValue
        Case "gameover"
            gameOver
            Exit Sub
        Case "gamewin"
'            gameWin
            Exit Sub
        Case "quit"
            Exit Sub
    End Select
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
    ActiveDocument.MasterPage.DesktopLayer.Visible = False
    Set s = ActivePage.Layers.Item(2).CreateArtisticText(400, 170, gameMain.ScorePoint, , , "Arial", 84, cdrTrue, cdrFalse, cdrNoFontLine, cdrCenterAlignment)
    s.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
    ActiveDocument.ClearSelection
End Sub

Private Sub gameWin()
    Dim s As Shape
    ActiveDocument.Pages.Item(3).Activate
    Set s = ActivePage.Layers.Item(2).CreateArtisticText(400, 170, 0, , , "Arial", 84, cdrTrue, cdrFalse, cdrNoFontLine, cdrCenterAlignment)
    s.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
    ActiveDocument.ClearSelection
End Sub

Private Sub UpdateInput()
    If (GetAsyncKeyState(vbKeySpace)) Then
          imDone = True
    End If
End Sub
