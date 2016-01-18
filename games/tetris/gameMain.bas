Attribute VB_Name = "gameMain"
Option Explicit

Dim imDone As Boolean
Dim tmr As Double
Dim cellSize As Integer
Dim boardHeight As Integer, boardWidth As Integer
Dim screenWidth As Integer, screenHeight As Integer
Dim Tick As Double
Dim keyReadDone As Boolean
Dim returnValue As String
Dim gc As New Collection

Sub LoadLevel()
    Tick = 0.3
    cellSize = 15
    boardHeight = 20
    boardWidth = 10
    scorePoint = 0
End Sub

Function GameLoop() As String
    Initialization
    LoadResource
    tmr = Timer
    Do Until imDone
        DoEvents
        UpdateInput
        If Timer > tmr + Tick And Not imDone Then
            Update
            Draw
            tmr = Timer
        End If
    Loop
    Destroy
    GameLoop = returnValue
End Function

Private Sub LoadResource()
    Dim sr As New ShapeRange
    Dim s As Shape
    Set gc = New Collection
    
End Sub

Private Sub Initialization()
    imDone = False
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
            s.Outline.Color.CMYKAssign 0, 0, 0, 20
            s.Outline.width = 0.1
        Next e
    Next i
    ActiveDocument.ClearSelection
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
End Sub

Private Sub Update()
    keyReadDone = False
    
End Sub

Private Sub Draw()
    Application.Optimization = True

    
    ActivePage.Layers.Item(2).Shapes.All.Delete
    ActivePage.Layers.Item(3).Shapes.All.Delete
    

    
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

