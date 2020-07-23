Attribute VB_Name = "PrintMarksR7"
Sub PrintMarksR7()
' version 1.2.1
' remove command group
' maybe this call bug

' version 1.2
' add create color bar in new document
If (Documents.Count = 0) Then Exit Sub

' maybe this call bug
' ActiveDocument.BeginCommandGroup "Create R7 Print Marks"
Application.Optimization = True
ActiveDocument.Unit = cdrMillimeter

    Dim iCB As Shape
    Dim colorBar As ShapeRange, cbLeftPart As ShapeRange, cbRightPart As ShapeRange
    Dim cbCrop As New ShapeRange, srD
    Dim cbTopPart As ShapeRange, cbBottomPart As ShapeRange
    Dim leftOffsetMark As ShapeRange
    Dim rightOffsetMark As ShapeRange
    Dim leftTargetMark As ShapeRange
    Dim rightTargetMark As ShapeRange
    Dim leftMark As ShapeRange
    Dim signCmyk As ShapeRange
    Dim printMarksPath As String
    Dim offsetLeftMark As Integer, offsetTargetMark As Integer, offsetColorBar As Integer, offsetBothSide
    Dim offsetSignColor As Integer
    Dim allMarks As ShapeRange
    Dim i As Integer, a As Integer
    Dim cBar As Shape, cbT As Shape, cbB As Shape
    
    'create color bar in new document
    Dim mainDoc As Document
    Dim mainPage As Page
    Dim cbarDoc As Document
    Set mainDoc = ActiveDocument
    Set mainPage = mainDoc.ActivePage
    Set cbarDoc = CreateDocument
    
    cbarDoc.Unit = cdrMillimeter
    cbarDoc.ActivePage.SizeWidth = mainPage.SizeWidth
    cbarDoc.ActivePage.SizeHeight = mainPage.SizeHeight
    
    printMarksPath = (UserDataPath & "printMarks\")
    offsetLeftMark = 55
    offsetTargetMark = 30
    offsetSignColor = 45
    offsetColorBar = 2
    offsetBothSide = 5
    
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
    signCmyk.PositionY = ActivePage.BoundingBox.Bottom + offsetSignColor
    
    colorBar.PositionY = ActivePage.BoundingBox.Bottom + colorBar.BoundingBox.Height + offsetColorBar
    cbTopPart.PositionY = colorBar.BoundingBox.Top + cbTopPart.BoundingBox.Height
    cbBottomPart.PositionY = colorBar.BoundingBox.Bottom
    ActiveDocument.ClearSelection
       
    cbTopPart.Ungroup
    cbTopPart.Item(1).Delete
    Set cbTopPart = ActiveSelectionRange
    ActiveDocument.ClearSelection
    
    cbBottomPart.Ungroup
    cbBottomPart.Item(1).Delete
    Set cbBottomPart = ActiveSelectionRange
    ActiveDocument.ClearSelection
    
    colorBar.Ungroup
    ActiveDocument.ClearSelection
    
    Set iCB = colorBar.Item(1)
    iCB.Ungroup
    Set cbLeftPart = ActiveSelectionRange
    ActiveDocument.ClearSelection
    
    Set iCB = colorBar.Item(2)
    iCB.Ungroup
    Set cbRightPart = ActiveSelectionRange
    ActiveDocument.ClearSelection
    
    '\ cut on a page
    Set cbCrop = New ShapeRange
    For Each iCB In cbLeftPart
        If iCB.BoundingBox.Left > ActivePage.BoundingBox.Left + offsetBothSide Then cbCrop.Add iCB
    Next iCB
    Set srD = cbCrop.Duplicate
    cbLeftPart.Delete
    Set cbLeftPart = srD
    
    Set cbCrop = New ShapeRange
    For Each iCB In cbRightPart
        If iCB.BoundingBox.Right < ActivePage.BoundingBox.Right - offsetBothSide Then cbCrop.Add iCB
    Next iCB
    Set srD = cbCrop.Duplicate
    cbRightPart.Delete
    Set cbRightPart = srD
    
    '\ cut on a condition
    Set cbCrop = New ShapeRange
    For i = 1 To cbLeftPart.Count
        If nextItem(cbLeftPart, i) Then
            For a = i To cbLeftPart.Count
                cbCrop.Add cbLeftPart.Item(a)
            Next a
            Exit For
        End If
    Next i
    Set srD = cbCrop.Duplicate
    cbLeftPart.Delete
    Set cbLeftPart = srD
    
    Set cbCrop = New ShapeRange
    For i = 1 To cbRightPart.Count
        If nextItem(cbRightPart, i) Then
            For a = i To cbRightPart.Count
                cbCrop.Add cbRightPart.Item(a)
            Next a
            Exit For
        End If
    Next i
    Set srD = cbCrop.Duplicate
    cbRightPart.Delete
    Set cbRightPart = srD
    
    Set colorBar = New ShapeRange
    colorBar.AddRange cbLeftPart
    colorBar.AddRange cbRightPart
    Set cBar = colorBar.Group
    
    '\ cut top and bottom part
    Set cbCrop = New ShapeRange
    For Each iCB In cbTopPart
        If (iCB.BoundingBox.Left > colorBar.BoundingBox.Left) And (iCB.BoundingBox.Right < colorBar.BoundingBox.Right) Then
            cbCrop.Add iCB
        End If
    Next iCB
    Set cbT = cbCrop.Duplicate.Group
    cbTopPart.Delete
    
    Set cbCrop = New ShapeRange
    For Each iCB In cbBottomPart
        If (iCB.BoundingBox.Left > colorBar.BoundingBox.Left) And (iCB.BoundingBox.Right < colorBar.BoundingBox.Right) Then
            cbCrop.Add iCB
        End If
    Next iCB
    Set cbB = cbCrop.Duplicate.Group
    cbBottomPart.Delete
    '\
    
    ActiveDocument.ClearSelection
    Set colorBar = New ShapeRange
    colorBar.Add cBar
    colorBar.Add cbT
    colorBar.Add cbB
    Set cBar = colorBar.Group
    
    ActiveDocument.ClearSelection
    cBar.AddToSelection
    leftOffsetMark.AddToSelection
    rightOffsetMark.AddToSelection
    leftMark.AddToSelection
    leftTargetMark.AddToSelection
    rightTargetMark.AddToSelection
    signCmyk.AddToSelection
    ActiveSelectionRange.Group.Copy
    
    mainPage.ActiveLayer.Paste
    cbarDoc.Dirty = False
    cbarDoc.Close
    
    mainDoc.Activate
    mainPage.Activate
    mainDoc.ClearSelection
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh

' ActiveDocument.EndCommandGroup
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



