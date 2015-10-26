Attribute VB_Name = "groupAll"
Sub GroupAll()
If (Documents.Count > 0) Then
    Dim aLayer As Layer
    Dim s1 As Shape
    Dim aPage As Page
    Dim guideL As Boolean
    Set aPage = ActivePage
    guideL = False
    If aPage.GuidesLayer.Editable Then
        guideL = True
        aPage.GuidesLayer.Editable = False
        aPage.GuidesLayer.Printable = False
    End If
    For Each aLayer In aPage.Layers
        If aLayer.Editable Then
            aLayer.Activate
            aLayer.Shapes.All.Group
        End If
    Next aLayer
    aPage.GuidesLayer.Editable = guideL
End If
End Sub
