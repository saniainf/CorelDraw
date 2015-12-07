Attribute VB_Name = "ShapeWeld"
Sub ShapeWeld()
    Application.Optimization = True
    Dim s1 As Shape
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    Set s1 = sr.FirstShape.Duplicate
    For i = 2 To sr.Count
        Dim s2 As Shape
        Set s2 = s1.Weld(sr.Item(i), False, False)
        Set s1 = s2
    Next i
    sr.FirstShape.Delete
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub
