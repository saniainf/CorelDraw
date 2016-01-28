Attribute VB_Name = "mMain"
Option Explicit
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim imDone As Boolean

Sub PipeGame()
    Application.ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft
    Dim returnValue As String
    
'    newGame
    ActiveDocument.Pages.Item(1).Activate
    gameMain.LoadLevel
    returnValue = gameMain.GameLoop
    Select Case returnValue
        Case "gameover"
'            gameOver
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

End Sub

Private Sub gameWin()

End Sub

Private Sub UpdateInput()
    If (GetAsyncKeyState(vbKeySpace)) Then
          imDone = True
    End If
End Sub
