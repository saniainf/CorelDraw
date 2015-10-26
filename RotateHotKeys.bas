Attribute VB_Name = "RotateHotKeys"
Sub RotateClockWise()
If (Documents.Count > 0) Then
    Dim activeSelection As ShapeRange
    If (ActiveSelectionRange.Count > 0) Then
        Set activeSelection = ActiveSelectionRange
        activeSelection.Rotate (-90)
    End If
End If
End Sub

Sub RotateCounterClockWise()
If (Documents.Count > 0) Then
    Dim activeSelection As ShapeRange
    If (ActiveSelectionRange.Count > 0) Then
        Set activeSelection = ActiveSelectionRange
        activeSelection.Rotate (90)
    End If
End If
End Sub
