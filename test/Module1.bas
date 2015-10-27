Attribute VB_Name = "Module1"
Sub Module1()
If (Documents.Count = 0) Then Exit Sub
    ActiveDocument.Unit = cdrMillimeter
    Application.Optimization = True
    
    Dim iCB As Shape
    Dim colorBar As ShapeRange, cbLeftPart As ShapeRange, cbRightPart As ShapeRange
    Dim cbCrop As New ShapeRange
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
    Dim i As Integer, a As Integer
    
    printMarksPath = ("e:\Projects\Scripts\CorelDraw\cdrFiles\printMarks\")
    offsetLeftMark = 55
    offsetTargetMark = 15
    offsetColorBar = 2
    
    ActiveLayer.Import (printMarksPath & "colorBarR7BodyPart.cdr")
    Set colorBar = ActiveSelectionRange
  
    colorBar.PositionY = ActivePage.BoundingBox.Bottom + colorBar.BoundingBox.Height + offsetColorBar
    colorBar.Ungroup

    Set iCB = colorBar.Item(1)
    iCB.Ungroup
    Set cbLeftPart = ActiveSelectionRange
    ActiveDocument.ClearSelection
    
    Set iCB = colorBar.Item(2)
    iCB.Ungroup
    Set cbRightPart = ActiveSelectionRange
    ActiveDocument.ClearSelection
    '\
    For Each iCB In cbLeftPart
        If iCB.BoundingBox.Left > ActivePage.BoundingBox.Left Then cbCrop.Add iCB
    Next iCB
    Set cbCrop = cbCrop.Duplicate
    cbLeftPart.Delete
    Set cbLeftPart = cbCrop
    Set cbCrop = New ShapeRange
    
    For Each iCB In cbRightPart
        If iCB.BoundingBox.Right < ActivePage.BoundingBox.Right Then cbCrop.Add iCB
    Next iCB
    Set cbCrop = cbCrop.Duplicate
    cbRightPart.Delete
    Set cbRightPart = cbCrop
    Set cbCrop = New ShapeRange
    '\
    For i = 1 To cbLeftPart.Count
        If nextItem(cbLeftPart, i) Then
            For a = i To cbLeftPart.Count
                cbCrop.Add cbLeftPart.Item(a)
            Next a
            Exit For
        End If
    Next i
    Set cbCrop = cbCrop.Duplicate
    cbLeftPart.Delete
    Set cbLeftPart = cbCrop
    Set cbCrop = New ShapeRange
    
    For i = 1 To cbRightPart.Count
        If nextItem(cbRightPart, i) Then
            For a = i To cbRightPart.Count
                cbCrop.Add cbRightPart.Item(a)
            Next a
            Exit For
        End If
    Next i
    Set cbCrop = cbCrop.Duplicate
    cbRightPart.Delete
    Set cbRightPart = cbCrop
    Set cbCrop = New ShapeRange
    '\
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Public Function nextItem(aSel As ShapeRange, i As Integer) As Boolean
    nextItem = False
    If i + 2 > aSel.Count Then
        nextItem = False
    ElseIf ((fillCmyk(aSel.Item(i))) And (fillCmyk(aSel.Item(i + 1)))) Then
        nextItem = True
    End If
End Function

Public Function fillCmyk(s1 As Shape) As Boolean
    Dim cyanColor As New Color
    Dim magentaColor As New Color
    Dim yellowColor As New Color
    Dim blackColor As New Color
    
    cyanColor.CMYKAssign 100, 0, 0, 0
    magentaColor.CMYKAssign 0, 100, 0, 0
    yellowColor.CMYKAssign 0, 0, 100, 0
    blackColor.CMYKAssign 0, 0, 0, 100

    fillCmyk = False
    If s1.Fill.UniformColor.IsSame(cyanColor) Then fillCmyk = True
    If s1.Fill.UniformColor.IsSame(magentaColor) Then fillCmyk = True
    If s1.Fill.UniformColor.IsSame(yellowColor) Then fillCmyk = True
    If s1.Fill.UniformColor.IsSame(blackColor) Then fillCmyk = True
End Function


