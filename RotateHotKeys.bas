Attribute VB_Name = "RotateHotKeys"
Sub RotateClockWise()
If (Documents.Count = 0) Then Exit Sub
If (ActiveSelectionRange.Count <= 0) Then Exit Sub
    Dim activeSelection As ShapeRange
    Set activeSelection = ActiveSelectionRange
    activeSelection.Rotate (-90)
End Sub

Sub RotateCounterClockWise()
If (Documents.Count = 0) Then Exit Sub
If (ActiveSelectionRange.Count <= 0) Then Exit Sub
    Dim activeSelection As ShapeRange
    Set activeSelection = ActiveSelectionRange
    activeSelection.Rotate (90)
End Sub
