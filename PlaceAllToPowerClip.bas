Attribute VB_Name = "PlaceAllToPowerClip"
Sub PlaceAllToPowerClip()
If (Documents.Count = 0) Then Exit Sub
ActiveDocument.BeginCommandGroup "Place to Power Clip"
Application.Optimization = True
    Application.ActiveDocument.Unit = cdrMillimeter
    Dim aSel As ShapeRange
    Dim shPowerClip As Shape
    Dim sL As Integer
    Dim sT As Integer
    Dim sR As Integer
    Dim sB As Integer
    Dim aPage As Page
    Dim aLayer As Layer
    Dim guideL As Boolean
    guideL = False
    Set aPage = ActivePage
    If aPage.GuidesLayer.Editable Then
        guideL = True
        aPage.GuidesLayer.Editable = False
        aPage.GuidesLayer.Printable = False
    End If
    sL = aPage.BoundingBox.Left
    sT = aPage.BoundingBox.Top
    sR = aPage.BoundingBox.Right
    sB = aPage.BoundingBox.Bottom
    For Each aLayer In aPage.Layers
        If aLayer.Editable Then
            If aLayer.Shapes.All.Count > 0 Then
                Set aSel = aLayer.Shapes.All
                Set shPowerClip = aLayer.CreateRectangle(sL, sT, sR, sB)
                shPowerClip.Outline.SetNoOutline
                shPowerClip.Fill.ApplyNoFill
                aSel.AddToPowerClip shPowerClip, cdrFalse
            End If
        End If
    Next aLayer
    aPage.GuidesLayer.Editable = guideL
    ActiveDocument.ClearSelection
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh
ActiveDocument.EndCommandGroup
End Sub
