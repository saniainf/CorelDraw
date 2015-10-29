Attribute VB_Name = "groupAll"
Sub GroupAll()
If (Documents.Count = 0) Then Exit Sub
    Application.Optimization = True
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
            If aLayer.Shapes.All.Count > 1 Then aLayer.Shapes.All.Group
        End If
    Next aLayer
    aPage.GuidesLayer.Editable = guideL
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub
