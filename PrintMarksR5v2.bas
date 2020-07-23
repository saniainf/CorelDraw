Attribute VB_Name = "PrintMarksR5v2"
Sub showForm()
    If (Documents.Count = 0) Then Exit Sub
    PrintMarksR5.Show vbModeles
End Sub

Public Sub PrintMarksR5v2(sColorBar As Shape, signColor As Shape)
    Dim itemColorBar As Shape
    Dim colorBar As New ShapeRange, finalColorBar As New ShapeRange, srD As New ShapeRange
    Dim leftOffsetMark As New ShapeRange
    Dim rightOffsetMark As New ShapeRange
    Dim leftTargetMark As New ShapeRange
    Dim rightTargetMark As New ShapeRange
    Dim leftMark As New ShapeRange
    Dim printMarksPath As String
    Dim offsetLeftMark As Integer, offsetTargetMark As Integer
    Dim allMarks As New ShapeRange
    
    printMarksPath = (UserDataPath & "printMarks\")
    offsetLeftMark = 55
    offsetTargetMark = 30
    
    colorBar.Add sColorBar
    
    ActiveLayer.Import (printMarksPath & "leftOffsetMark.cdr")
    Set leftOffsetMark = ActiveSelectionRange
    ActiveLayer.Import (printMarksPath & "rightOffsetMark.cdr")
    Set rightOffsetMark = ActiveSelectionRange
    ActiveLayer.Import (printMarksPath & "targetMark.cdr")
    Set leftTargetMark = ActiveSelectionRange
    Set rightTargetMark = leftTargetMark.Duplicate
    ActiveLayer.Import (printMarksPath & "leftMark.cdr")
    Set leftMark = ActiveSelectionRange
    Set finalColorBar = New ShapeRange

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
    colorBar.CenterX = ActivePage.BoundingBox.CenterX
    colorBar.PositionY = ActivePage.BoundingBox.Bottom + colorBar.BoundingBox.Height
    signColor.Rotate 90
    signColor.PositionX = leftTargetMark.BoundingBox.CenterX - signColor.BoundingBox.Width / 2
    signColor.PositionY = leftTargetMark.BoundingBox.Bottom + signColor.BoundingBox.Height + offsetTargetMark
    
    colorBar.Ungroup
    For Each itemColorBar In colorBar
        If (itemColorBar.BoundingBox.Left > ActivePage.BoundingBox.Left) And (itemColorBar.BoundingBox.Right < ActivePage.BoundingBox.Right) Then
            finalColorBar.Add itemColorBar
        End If
    Next itemColorBar
    Set srD = finalColorBar.Duplicate
    colorBar.Delete
    
    Set itemColorBar = srD.Group
    itemColorBar.AddToSelection
    leftOffsetMark.AddToSelection
    rightOffsetMark.AddToSelection
    leftMark.AddToSelection
    leftTargetMark.AddToSelection
    rightTargetMark.AddToSelection
    signColor.AddToSelection
    ActiveSelectionRange.Group.Copy
End Sub

