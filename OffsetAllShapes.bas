Attribute VB_Name = "OffsetAllShapes"
Sub OffsetAllShapes12()
If Documents.Count > 0 Then
    Dim guideL As Boolean
    Dim masterL As Boolean
    guideL = False
    desktopL = False
    Application.ActiveDocument.Unit = cdrMillimeter
    If ActiveDocument.MasterPage.GuidesLayer.Editable Then
        guideL = True
        ActiveDocument.MasterPage.GuidesLayer.Editable = False
    End If
    If ActiveDocument.MasterPage.DesktopLayer.Editable Then
        desktopL = True
        ActiveDocument.MasterPage.DesktopLayer.Editable = False
    End If
    ActivePage.Shapes.All.CreateSelection
    activeSelection.TopY = ActivePage.BoundingBox.Top - 30
    ActiveDocument.MasterPage.GuidesLayer.Editable = guideL
    ActiveDocument.MasterPage.DesktopLayer.Editable = desktopL
End If
End Sub

Sub OffsetAllShapes10()
If Documents.Count > 0 Then
    Dim guideL As Boolean
    guideL = False
    desktopL = False
    Application.ActiveDocument.Unit = cdrMillimeter
    If ActiveDocument.MasterPage.GuidesLayer.Editable Then
        guideL = True
        ActiveDocument.MasterPage.GuidesLayer.Editable = False
    End If
    If ActiveDocument.MasterPage.DesktopLayer.Editable Then
        desktopL = True
        ActiveDocument.MasterPage.DesktopLayer.Editable = False
    End If
    ActivePage.Shapes.All.CreateSelection
    activeSelection.TopY = ActivePage.BoundingBox.Top - 28
    ActiveDocument.MasterPage.GuidesLayer.Editable = guideL
    ActiveDocument.MasterPage.DesktopLayer.Editable = desktopL
End If
End Sub
