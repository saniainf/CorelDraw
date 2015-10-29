Attribute VB_Name = "DistributeToEdge"
Sub DistributeLeftEdge()
If (Documents.Count = 0) Then Exit Sub
If (ActiveSelectionRange.Count < 2) Then Exit Sub
Application.ActiveDocument.Unit = cdrMillimeter
    Dim activeSelection As ShapeRange
    Dim selectShape As Shape
    Dim aX As Double
    Dim aY As Double
    Dim aW As Double
    Dim aH As Double 'anchor shape
    Dim sX As Double
    Dim sY As Double
    Dim sW As Double
    Dim sH As Double 'select shape
    Dim firstShape As Boolean
    
    If (ActiveSelectionRange.Count >= 2) Then
        Set activeSelection = ActiveSelectionRange
        firstShape = True
        For Each selectShape In activeSelection
            If (firstShape) Then
                aX = selectShape.PositionX
                aY = selectShape.PositionY
                aW = selectShape.SizeWidth
                aH = selectShape.SizeHeight
            Else
                sX = selectShape.PositionX
                sY = selectShape.PositionY
                sW = selectShape.SizeWidth
                sH = selectShape.SizeHeight
                selectShape.PositionX = aX - sW
                aX = selectShape.PositionX
                aY = selectShape.PositionY
                aW = selectShape.SizeWidth
                aH = selectShape.SizeHeight
            End If
            firstShape = False
        Next selectShape
    End If
End Sub

Sub DistributeRightEdge()
If (Documents.Count = 0) Then Exit Sub
If (ActiveSelectionRange.Count < 2) Then Exit Sub
Application.ActiveDocument.Unit = cdrMillimeter
    Dim activeSelection As ShapeRange
    Dim selectShape As Shape
    Dim aX As Double
    Dim aY As Double
    Dim aW As Double
    Dim aH As Double 'anchor shape
    Dim sX As Double
    Dim sY As Double
    Dim sW As Double
    Dim sH As Double 'select shape
    Dim firstShape As Boolean
    
    If (ActiveSelectionRange.Count >= 2) Then
        Set activeSelection = ActiveSelectionRange
        firstShape = True
        For Each selectShape In activeSelection
            If (firstShape) Then
                aX = selectShape.PositionX
                aY = selectShape.PositionY
                aW = selectShape.SizeWidth
                aH = selectShape.SizeHeight
            Else
                sX = selectShape.PositionX
                sY = selectShape.PositionY
                sW = selectShape.SizeWidth
                sH = selectShape.SizeHeight
                selectShape.PositionX = aX + aW
                aX = selectShape.PositionX
                aY = selectShape.PositionY
                aW = selectShape.SizeWidth
                aH = selectShape.SizeHeight
            End If
            firstShape = False
        Next selectShape
    End If
End Sub

Sub DistributeTopEdge()
If (Documents.Count = 0) Then Exit Sub
If (ActiveSelectionRange.Count < 2) Then Exit Sub
Application.ActiveDocument.Unit = cdrMillimeter
    Dim activeSelection As ShapeRange
    Dim selectShape As Shape
    Dim aX As Double
    Dim aY As Double
    Dim aW As Double
    Dim aH As Double 'anchor shape
    Dim sX As Double
    Dim sY As Double
    Dim sW As Double
    Dim sH As Double 'select shape
    Dim firstShape As Boolean
    
    If (ActiveSelectionRange.Count >= 2) Then
        Set activeSelection = ActiveSelectionRange
        firstShape = True
        For Each selectShape In activeSelection
            If (firstShape) Then
                aX = selectShape.PositionX
                aY = selectShape.PositionY
                aW = selectShape.SizeWidth
                aH = selectShape.SizeHeight
            Else
                sX = selectShape.PositionX
                sY = selectShape.PositionY
                sW = selectShape.SizeWidth
                sH = selectShape.SizeHeight
                selectShape.PositionY = aY + sH
                aX = selectShape.PositionX
                aY = selectShape.PositionY
                aW = selectShape.SizeWidth
                aH = selectShape.SizeHeight
            End If
            firstShape = False
        Next selectShape
    End If
End Sub

Sub DistributeBottomEdge()
If (Documents.Count = 0) Then Exit Sub
If (ActiveSelectionRange.Count < 2) Then Exit Sub
Application.ActiveDocument.Unit = cdrMillimeter
    Dim activeSelection As ShapeRange
    Dim selectShape As Shape
    Dim aX As Double
    Dim aY As Double
    Dim aW As Double
    Dim aH As Double 'anchor shape
    Dim sX As Double
    Dim sY As Double
    Dim sW As Double
    Dim sH As Double 'select shape
    Dim firstShape As Boolean
    
    If (ActiveSelectionRange.Count >= 2) Then
        Set activeSelection = ActiveSelectionRange
        firstShape = True
        For Each selectShape In activeSelection
            If (firstShape) Then
                aX = selectShape.PositionX
                aY = selectShape.PositionY
                aW = selectShape.SizeWidth
                aH = selectShape.SizeHeight
            Else
                sX = selectShape.PositionX
                sY = selectShape.PositionY
                sW = selectShape.SizeWidth
                sH = selectShape.SizeHeight
                selectShape.PositionY = aY - aH
                aX = selectShape.PositionX
                aY = selectShape.PositionY
                aW = selectShape.SizeWidth
                aH = selectShape.SizeHeight
            End If
            firstShape = False
        Next selectShape
    End If
End Sub
