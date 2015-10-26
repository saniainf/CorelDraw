Attribute VB_Name = "PrintGridOff"
Sub PrintGridOff()
If (Documents.Count > 0) Then
    '/
    With ActiveDocument.MasterPage.GridLayer
        .Visible = False
        .Editable = False
        .Printable = False
    End With
     
    With ActiveDocument.MasterPage.DesktopLayer
        .Printable = False
    End With
     
    With ActiveDocument.MasterPage.GuidesLayer
        .Printable = False
    End With
    '/
    Dim aPage As Page
    For Each aPage In ActiveDocument.Pages
        aPage.GuidesLayer.Printable = False
    Next aPage
    ActiveDocument.Pages(1).Activate
End If
End Sub


