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
Public Enum pieceTypes
    Sunday = 1
    Monday = 2
    Tuesday = 3
    Wednesday = 4
    Thursday = 5
    Friday = 6
    Saturday = 7
End Enum

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
        Update
        Draw
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
    
    drawGameField
    
    testsub
End Sub

Private Sub testsub()
    Dim i As Integer, e As Integer
    Dim s As Shape
    Randomize
    Application.Optimization = True
    For i = 0 To boardHeight - 1
        For e = 0 To boardWidth - 1
            Set s = gcWaterPipe.Item(Int((6 * Rnd) + 1)).Duplicate
            s.MoveToLayer ActivePage.Layers.Item(2)
            s.SetPosition e * cellSize, i * cellSize
        Next e
    Next i
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
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
        drawBoard = True
        keyReadDone = True
    ElseIf (GetAsyncKeyState(vbKeyDown)) And Not keyReadDone Then
        
        keyReadDone = True
    ElseIf (GetAsyncKeyState(vbKeyLeft)) And Not keyReadDone Then
        
        keyReadDone = True
    ElseIf (GetAsyncKeyState(vbKeyRight)) And Not keyReadDone Then
        
        keyReadDone = True
    End If
End Sub

Private Sub Update()
    keyReadDone = False
    
End Sub

Private Sub Draw()
    If drawBoard Then
        Application.Optimization = True
        ActivePage.Layers.Item(2).Shapes.All.Delete
        ActivePage.Layers.Item(3).Shapes.All.Delete

    
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
