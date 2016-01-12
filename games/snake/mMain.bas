Attribute VB_Name = "mMain"
Sub SnakeGame()
    Application.ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrTopLeft
    Dim i As Integer
    
    For i = 1 To gameLevel.levelCount
        ActiveDocument.Pages.Item(1).Activate
        gameLevel.levelChange (i)
        gameMain.LoadLevel
        If gameMain.GameLoop = 0 Then
            gameOver
            Exit Sub
        End If
        If gameMain.GameLoop = 1 Then
            nextLevel
        End If
    Next i
    gameWin
End Sub

Private Sub gameOver()
    MsgBox "Game Over"
End Sub

Private Sub nextLevel()

End Sub

Private Sub gameWin()
    MsgBox "Game Win"
End Sub
