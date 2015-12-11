Attribute VB_Name = "DuplicateHotKeys"
Sub DuplicateTop()
If (Documents.Count = 0) Then Exit Sub
If (ActiveSelectionRange.Count < 1) Then Exit Sub
ActiveDocument.BeginCommandGroup "Duplicate shape"
Application.Optimization = True
Application.ActiveDocument.Unit = cdrMillimeter
    Dim activeSelection As ShapeRange
    Dim dShape As ShapeRange
    Dim sX As Double
    Dim sY As Double
    Dim sW As Double
    Dim sH As Double
    If (ActiveSelectionRange.Count > 0) Then
        Set activeSelection = ActiveSelectionRange
        sX = activeSelection.PositionX
        sY = activeSelection.PositionY
        sW = activeSelection.SizeWidth
        sH = activeSelection.SizeHeight
        Set dShape = activeSelection.Duplicate
        dShape.Move 0#, sH
        dShape.CreateSelection
    End If
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh
ActiveDocument.EndCommandGroup
End Sub

Sub DuplicateBottom()
If (Documents.Count = 0) Then Exit Sub
If (ActiveSelectionRange.Count < 1) Then Exit Sub
ActiveDocument.BeginCommandGroup "Duplicate shape"
Application.Optimization = True
Application.ActiveDocument.Unit = cdrMillimeter
    Dim activeSelection As ShapeRange
    Dim dShape As ShapeRange
    Dim sX As Double
    Dim sY As Double
    Dim sW As Double
    Dim sH As Double
    If (ActiveSelectionRange.Count > 0) Then
        Set activeSelection = ActiveSelectionRange
        sX = activeSelection.PositionX
        sY = activeSelection.PositionY
        sW = activeSelection.SizeWidth
        sH = activeSelection.SizeHeight
        Set dShape = activeSelection.Duplicate
        dShape.Move 0#, -sH
        dShape.CreateSelection
    End If
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh
ActiveDocument.EndCommandGroup
End Sub

Sub DuplicateLeft()
If (Documents.Count = 0) Then Exit Sub
If (ActiveSelectionRange.Count < 1) Then Exit Sub
ActiveDocument.BeginCommandGroup "Duplicate shape"
Application.Optimization = True
Application.ActiveDocument.Unit = cdrMillimeter
    Dim activeSelection As ShapeRange
    Dim dShape As ShapeRange
    Dim sX As Double
    Dim sY As Double
    Dim sW As Double
    Dim sH As Double
    If (ActiveSelectionRange.Count > 0) Then
        Set activeSelection = ActiveSelectionRange
        sX = activeSelection.PositionX
        sY = activeSelection.PositionY
        sW = activeSelection.SizeWidth
        sH = activeSelection.SizeHeight
        Set dShape = activeSelection.Duplicate
        dShape.Move -sW, 0#
        dShape.CreateSelection
    End If
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh
ActiveDocument.EndCommandGroup
End Sub

Sub DuplicateRight()
If (Documents.Count = 0) Then Exit Sub
If (ActiveSelectionRange.Count < 1) Then Exit Sub
ActiveDocument.BeginCommandGroup "Duplicate shape"
Application.Optimization = True
Application.ActiveDocument.Unit = cdrMillimeter
    Dim activeSelection As ShapeRange
    Dim dShape As ShapeRange
    Dim sX As Double
    Dim sY As Double
    Dim sW As Double
    Dim sH As Double
    If (ActiveSelectionRange.Count > 0) Then
        Set activeSelection = ActiveSelectionRange
        sX = activeSelection.PositionX
        sY = activeSelection.PositionY
        sW = activeSelection.SizeWidth
        sH = activeSelection.SizeHeight
        Set dShape = activeSelection.Duplicate
        dShape.Move sW, 0#
        dShape.CreateSelection
    End If
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh
ActiveDocument.EndCommandGroup
End Sub

