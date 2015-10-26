Attribute VB_Name = "RoundShapeSizePosition"
Sub RoundShapeSizePosition()
If (Documents.Count > 0) Then
    Application.ActiveDocument.Unit = cdrMillimeter
    Dim aSelection As ShapeRange
    Set aSelection = ActiveSelectionRange
    aSelection.SizeHeight = Math.Round(aSelection.SizeHeight)
    aSelection.SizeWidth = Math.Round(aSelection.SizeWidth)
    aSelection.PositionX = Math.Round(aSelection.PositionX, 0)
    aSelection.PositionY = Math.Round(aSelection.PositionY, 0)
End If
End Sub
