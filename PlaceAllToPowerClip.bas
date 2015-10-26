Attribute VB_Name = "PlaceAllToPowerClip"
Sub PlaceAllToPowerClip()
If Documents.Count > 0 Then
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
            Set aSel = aLayer.Shapes.All
            Set shPowerClip = aLayer.CreateRectangle(sL, sT, sR, sB)
            shPowerClip.Outline.SetNoOutline
            aSel.AddToPowerClip shPowerClip, cdrFalse
        End If
    Next aLayer
    aPage.GuidesLayer.Editable = guideL
End If
End Sub
