Attribute VB_Name = "PlaceAllToPowerClip"
Sub PlaceAllToPowerClip()
If Documents.Count > 0 Then
    Application.ActiveDocument.Unit = cdrMillimeter
    '/
    With ActiveDocument.MasterPage.GridLayer
        .Visible = False
        .Editable = False
        .Printable = False
    End With
     
    With ActiveDocument.MasterPage.DesktopLayer
        .Visible = False
        .Editable = False
        .Printable = False
    End With
     
    With ActiveDocument.MasterPage.GuidesLayer
        .Editable = False
        .Printable = False
    End With
    '/
    Dim aSel As ShapeRange
    Dim shPowerClip As Shape
    Dim sL As Integer
    Dim sT As Integer
    Dim sR As Integer
    Dim sB As Integer
    Dim aPage As Page
    Dim aLayer As Layer
    For Each aPage In ActiveDocument.Pages
        aPage.GuidesLayer.Editable = False
        aPage.GuidesLayer.Printable = False
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
    Next aPage
    ActiveDocument.Pages(1).Activate
End If
End Sub
