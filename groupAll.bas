Attribute VB_Name = "groupAll"
Sub GroupAll()
If (Documents.Count > 0) Then
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
    Dim aPage As Page
    Dim aLayer As Layer
    Dim s1 As Shape
    For Each aPage In ActiveDocument.Pages
        aPage.Activate
        aPage.GuidesLayer.Editable = False
        aPage.GuidesLayer.Printable = False
        For Each aLayer In aPage.Layers
            If aLayer.Editable Then
                aLayer.Activate
                aLayer.Shapes.All.Group
            End If
        Next aLayer
    Next aPage
    ActiveDocument.Pages(1).Activate
End If
End Sub
