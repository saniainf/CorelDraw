Attribute VB_Name = "colorbar"
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
    fillCmyk = False
    If s1.Fill.UniformColor.HexValue = "#00A0E3" Then fillCmyk = True
    If s1.Fill.UniformColor.HexValue = "#E5097F" Then fillCmyk = True
    If s1.Fill.UniformColor.HexValue = "#FFED00" Then fillCmyk = True
    If s1.Fill.UniformColor.HexValue = "#2B2A29" Then fillCmyk = True
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
            aSel.Item(i).Delete
        End If
    Next i
End Sub

Public Function nextItem(aSel As ShapeRange, i As Integer) As Boolean
    nextItem = False
    If i + 2 > aSel.Count Then
        nextItem = False
    ElseIf ((fillCmyk(aSel.Item(i))) And (fillCmyk(aSel.Item(i + 1))) And (fillCmyk(aSel.Item(i + 2)))) Then
        nextItem = True
    End If
End Function
