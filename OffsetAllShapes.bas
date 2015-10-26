Attribute VB_Name = "OffsetAllShapes"
Sub OffsetAllShapes12()
If Documents.Count > 0 Then
    Application.ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.MasterPage.GuidesLayer.Editable = False
    ActivePage.Shapes.All.CreateSelection
    activeSelection.TopY = ActivePage.BoundingBox.Top - 30
End If
End Sub

Sub OffsetAllShapes10()
If Documents.Count > 0 Then
    Application.ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.MasterPage.GuidesLayer.Editable = False
    ActivePage.Shapes.All.CreateSelection
    activeSelection.TopY = ActivePage.BoundingBox.Top - 28
End If
End Sub
