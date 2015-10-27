Attribute VB_Name = "PrintMarksR7"
Sub PrintMarksR7()
If (Documents.Count = 0) Then Exit Sub
    ActiveDocument.Unit = cdrMillimeter
    Application.Optimization = True

    Dim itemColorBar As Shape
    Dim colorBar As ShapeRange, cbLeftPart As ShapeRange, cbRightPart As ShapeRange
    Dim cbTopPart As ShapeRange, cbBottomPart As ShapeRange
    Dim leftOffsetMark As ShapeRange
    Dim rightOffsetMark As ShapeRange
    Dim leftTargetMark As ShapeRange
    Dim rightTargetMark As ShapeRange
    Dim leftMark As ShapeRange
    Dim signCmyk As ShapeRange
    Dim printMarksPath As String
    Dim offsetLeftMark As Integer, offsetTargetMark As Integer, offsetColorBar As Integer
    Dim allMarks As ShapeRange
    Dim i As Integer
    
    printMarksPath = ("e:\Projects\Scripts\CorelDraw\cdrFiles\printMarks\")
    offsetLeftMark = 55
    offsetTargetMark = 15
    offsetColorBar = 2
    
    ActiveLayer.Import (printMarksPath & "leftOffsetMark.cdr")
    Set leftOffsetMark = ActiveSelectionRange
    ActiveLayer.Import (printMarksPath & "rightOffsetMark.cdr")
    Set rightOffsetMark = ActiveSelectionRange
    ActiveLayer.Import (printMarksPath & "targetMark.cdr")
    Set leftTargetMark = ActiveSelectionRange
    Set rightTargetMark = leftTargetMark.Duplicate
    ActiveLayer.Import (printMarksPath & "leftMark.cdr")
    Set leftMark = ActiveSelectionRange
    ActiveLayer.Import (printMarksPath & "signCmyk.cdr")
    Set signCmyk = ActiveSelectionRange
    
    ActiveLayer.Import (printMarksPath & "colorBarR7BodyPart.cdr")
    Set colorBar = ActiveSelectionRange
    ActiveLayer.Import (printMarksPath & "colorBarR7TopPart.cdr")
    Set cbTopPart = ActiveSelectionRange
    ActiveLayer.Import (printMarksPath & "colorBarR7BottomPart.cdr")
    Set cbBottomPart = ActiveSelectionRange
    
    leftOffsetMark.PositionX = ActivePage.BoundingBox.Left
    leftOffsetMark.PositionY = ActivePage.BoundingBox.Top
    rightOffsetMark.PositionX = ActivePage.BoundingBox.Right - rightOffsetMark.BoundingBox.Width
    rightOffsetMark.PositionY = ActivePage.BoundingBox.Top
    leftMark.PositionX = ActivePage.BoundingBox.Left
    leftMark.PositionY = ActivePage.BoundingBox.Top - offsetLeftMark
    leftTargetMark.PositionX = ActivePage.BoundingBox.Left
    leftTargetMark.PositionY = ActivePage.BoundingBox.Bottom + offsetTargetMark
    rightTargetMark.PositionX = ActivePage.BoundingBox.Right - rightTargetMark.BoundingBox.Width
    rightTargetMark.PositionY = ActivePage.BoundingBox.Bottom + offsetTargetMark
    signCmyk.PositionX = leftTargetMark.BoundingBox.CenterX - signCmyk.BoundingBox.Width / 2
    signCmyk.PositionY = ActivePage.BoundingBox.Bottom + offsetTargetMark * 2
    
    colorBar.PositionY = ActivePage.BoundingBox.Bottom + colorBar.BoundingBox.Height + offsetColorBar
    cbTopPart.PositionY = colorBar.BoundingBox.Top + cbTopPart.BoundingBox.Height
    cbBottomPart.PositionY = colorBar.BoundingBox.Bottom
       
    cbTopPart.Ungroup
    cbTopPart.Item(1).Delete
    cbBottomPart.Ungroup
    cbBottomPart.Item(1).Delete
    colorBar.Ungroup
    'Set cbLeftPart = colorBar.Item(1).Shapes.All
    'Set cbRightPart = colorBar.Item(2).Shapes.All
    Set itemColorBar = colorBar.Item(1)
    itemColorBar.Ungroup
    Set cbLeftPart = ActiveSelectionRange
    Set itemColorBar = colorBar.Item(2)
    itemColorBar.Ungroup
    Set cbRightPart = ActiveSelectionRange
    
    For Each itemColorBar In cbLeftPart
        If itemColorBar.BoundingBox.Left < ActivePage.BoundingBox.Left Then itemColorBar.Delete
    Next itemColorBar
    
    For Each itemColorBar In cbRightPart
        If itemColorBar.BoundingBox.Right > ActivePage.BoundingBox.Right Then itemColorBar.Delete
    Next itemColorBar
    
    For i = 1 To cbLeftPart.Count
        If nextItem(cbLeftPart, i) Then
            Exit For
        Else
            cbLeftPart.Item(i).Delete
        End If
    Next i
    For i = 1 To cbRightPart.Count
        If nextItem(cbRightPart, i) Then
            Exit For
        Else
            cbRightPart.Item(i).Delete
        End If
    Next i
        
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh

End Sub

Public Function fillCmyk(s1 As Shape) As Boolean
    fillCmyk = False
    If s1.Fill.UniformColor.IsCMYK Then
        If s1.Fill.UniformColor.CMYKCyan = 100 Then fillCmyk = True
        If s1.Fill.UniformColor.CMYKMagenta = 100 Then fillCmyk = True
        If s1.Fill.UniformColor.CMYKYellow = 100 Then fillCmyk = True
        If s1.Fill.UniformColor.CMYKBlack = 100 Then fillCmyk = True
    End If
End Function

Public Function nextItem(aSel As ShapeRange, i As Integer) As Boolean
    nextItem = False
    If i + 2 > aSel.Count Then
        nextItem = False
    ElseIf ((fillCmyk(aSel.Item(i))) And (fillCmyk(aSel.Item(i + 1)))) Then
        nextItem = True
    End If
End Function

