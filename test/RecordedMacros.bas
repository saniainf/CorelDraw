Attribute VB_Name = "RecordedMacros"
Sub Macro5()
    Dim aSel As ShapeRange
    Dim s1 As Shape
    Dim cyanColor As New Color
    cyanColor.CMYKAssign 100, 0, 0, 0
    Set aSel = ActiveSelectionRange
    
    For Each s1 In aSel
        If (s1.Fill.UniformColor.IsSame(cyanColor)) Then
            s1.PositionY = s1.PositionY + 5
        End If
    Next s1
End Sub

Sub Macro1()
    Dim aSel As ShapeRange
    Dim g1 As Shape
    Dim i As Integer
    
    Application.Optimization = True
    ActiveLayer.Import ("d:\Projects\coreldraw\colorbar.cdr")
    Set aSel = ActiveSelectionRange
    aSel.Ungroup
    aSel.Ungroup
    For Each g1 In aSel
        If g1.PositionX < ActivePage.BoundingBox.Left Then
            g1.Delete
        ElseIf g1.PositionX + g1.SizeWidth > ActivePage.BoundingBox.Right Then
            g1.Delete
        End If
        i = i + 1
    Next g1
    aSel.Group
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
    MsgBox (i)
End Sub

Sub Macro2()
    Dim aSel As ShapeRange
    Dim g1 As Shape
    ActiveDocument.Unit = cdrMillimeter
    Set aSel = ActiveSelectionRange
    For Each g1 In aSel
        If fillCmyk(g1) Then
            g1.PositionY = g1.PositionY + g1.SizeHeight
            Exit For
        End If
    Next g1
End Sub

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
    If (s1.Fill.UniformColor.IsSame(cyanColor)) Then fillCmyk = True
    If s1.Fill.UniformColor.IsSame(magentaColor) Then fillCmyk = True
    If s1.Fill.UniformColor.IsSame(yellowColor) Then fillCmyk = True
    If s1.Fill.UniformColor.IsSame(blackColor) Then fillCmyk = True
End Function

Sub Macro3()
    Dim aSel As ShapeRange
    Dim i As Integer
    ActiveDocument.Unit = cdrMillimeter
    Set aSel = ActiveSelectionRange
    For i = 1 To aSel.Count
        If nextItem(aSel, i) Then
            Exit For
        Else
            aSel.Item(i).PositionY = aSel.Item(i).PositionY + 5
        End If
    Next i
End Sub

Public Function nextItem(aSel As ShapeRange, i As Integer) As Boolean
    nextItem = False
    If i + 2 > aSel.Count Then
        nextItem = False
    ElseIf ((fillCmyk(aSel.Item(i))) And (fillCmyk(aSel.Item(i + 1)))) Then
        nextItem = True
    End If
End Function
