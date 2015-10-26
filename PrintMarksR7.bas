Attribute VB_Name = "PrintMarksR7"
Sub PrintMarksR7()
If (Documents.Count > 0) Then
    ActiveDocument.Unit = cdrMillimeter
    Application.Optimization = True
    
    Dim itemColorBar As Shape
    Dim colorBar As ShapeRange
    Dim leftOffsetMark As ShapeRange
    Dim rightOffsetMark As ShapeRange
    Dim leftTargetMark As ShapeRange
    Dim rightTargetMark As ShapeRange
    Dim leftMark As ShapeRange
    Dim signCmyk As ShapeRange
    Dim printMarksPath As String
    Dim offsetLeftMark As Integer, offsetTargetMark As Integer
    Dim allMarks As ShapeRange
    
    printMarksPath = ("e:\Projects\Scripts\CorelDraw\cdrFiles\printMarks\")
    offsetLeftMark = 55
    offsetTargetMark = 15
    
    ActiveLayer.Import (printMarksPath & "colorBarR5.cdr")
    Set colorBar = ActiveSelectionRange
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
    
    colorBar.CenterX = ActivePage.BoundingBox.CenterX
    colorBar.PositionY = ActivePage.BoundingBox.Bottom + colorBar.BoundingBox.Height
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
    
    colorBar.Ungroup
    For Each itemColorBar In colorBar
        If itemColorBar.BoundingBox.Left < ActivePage.BoundingBox.Left Then
            itemColorBar.Delete
        ElseIf itemColorBar.BoundingBox.Right > ActivePage.BoundingBox.Right Then
            itemColorBar.Delete
        End If
    Next itemColorBar
    
    colorBar.AddToSelection
    leftOffsetMark.AddToSelection
    rightOffsetMark.AddToSelection
    leftMark.AddToSelection
    leftTargetMark.AddToSelection
    rightTargetMark.AddToSelection
    signCmyk.AddToSelection
    ActiveSelectionRange.Group
    
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End If
End Sub
