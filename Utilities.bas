Attribute VB_Name = "Utilities"
Sub OffLayerByName()
    Dim doc As Document
    Dim aPage As Page
    Dim l As Layer
    
    For Each doc In Application.Documents
        doc.Activate
        For Each aPage In doc.Pages
            For Each l In aPage.AllLayers
                If l.Name = "LAK" Then 'name
                    l.Visible = False
                    l.Printable = False
                End If
            Next l
        Next aPage
    Next doc
End Sub

Sub SignCityName()
Application.Optimization = True
    Dim doc As Document
    Dim aPage As Page
    Dim l As Layer
    Dim s1 As Shape
    Dim x As Double
    Dim y As Double
    
    For Each doc In Application.Documents
        doc.Activate
        doc.Unit = cdrMillimeter
            Set aPage = doc.Pages.First
            For Each l In aPage.AllLayers
'                l.Activate
'                l.Editable = True
'                l.SelectableShapes.All.Group
'                If l.Name = "Слой 1" Then
'                    x = aPage.LeftX + 14
'                    y = aPage.CenterY
'                    Set s1 = l.CreateArtisticText(x, y, "Томск", cdrLanguageNone, cdrCharSetMixed, "Arial", 9, cdrTrue, cdrFalse, cdrNoFontLine, cdrCenterAlignment)
'                    s1.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
'                    s1.Outline.SetNoOutline
'                    s1.Rotate 90
'                    l.SelectableShapes.All.Group
'                End If
                
                If l.Name = "stamp" Then
                    l.SelectableShapes.All.FirstShape.Outline.SetNoOutline
                    l.SelectableShapes.All.FirstShape.Fill.ApplyNoFill
                End If

                If l.Name = "ЛАК" Then
                    l.SelectableShapes.All.FirstShape.Outline.SetNoOutline
                    l.SelectableShapes.All.FirstShape.Fill.ApplyNoFill
'                    l.Printable = False
                End If
            Next l
    Next doc
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh
End Sub

Sub CloseAllDocWithSave()
        Dim doc As Document
        For Each doc In Application.Documents
            doc.Save
            doc.Close
        Next doc
End Sub

Sub CloseAllDocWithoutSave()
        Dim doc As Document
        For Each doc In Application.Documents
            ActiveDocument.Dirty = False
            doc.Close
        Next doc
End Sub

Sub PlaceAllToPowerClipOnAllDoc()
    Application.Optimization = True
    
        Dim doc As Document
        Dim aPage As Page
        For Each doc In Application.Documents
            doc.Activate
            doc.Unit = cdrMillimeter
            For Each aPage In doc.Pages
                aPage.Activate
                PlaceAllToPowerClip aPage
            Next aPage
        Next doc
        
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Sub PageSize()
    Application.Optimization = True
    
        Dim doc As Document
        Dim aPage As Page
        For Each doc In Application.Documents
            doc.Activate
            doc.Unit = cdrMillimeter
            For Each aPage In doc.Pages
                aPage.Activate
                aPage.SizeWidth = 150
                aPage.SizeHeight = 212
            Next aPage
        Next doc
    
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Sub PasteClipBoardOnAllDoc()
    If Clipboard.Empty And Clipboard.Valid Then
        MsgBox "There is no data in the Clipboard."
        Exit Sub
    End If
    Application.Optimization = True
    
        Dim doc As Document
        Dim aPage As Page
        Dim s As Shape
        For Each doc In Application.Documents
            doc.Activate
            doc.Unit = cdrMillimeter
            For Each aPage In doc.Pages
                aPage.Activate
                Set s = aPage.ActiveLayer.Paste
                s.OrderToBack 'order
            Next aPage
        Next doc
    
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Sub PasteOnAllPages()
    Dim p As Page
    For Each p In ActiveDocument.Pages
        p.Activate
        p.ActiveLayer.Paste
    Next p
End Sub

Sub PlaceAllToPowerClip(aPage As Page)
    Dim aSel As ShapeRange
    Dim shPowerClip As Shape
    Dim sL As Integer
    Dim sT As Integer
    Dim sr As Integer
    Dim sB As Integer
    Dim aLayer As Layer
    Dim guideL As Boolean
    guideL = False
    If aPage.GuidesLayer.Editable Then
        guideL = True
        aPage.GuidesLayer.Editable = False
        aPage.GuidesLayer.Printable = False
    End If
    sL = aPage.BoundingBox.Left
    sT = aPage.BoundingBox.Top
    sr = aPage.BoundingBox.Right
    sB = aPage.BoundingBox.Bottom
    For Each aLayer In aPage.Layers
        If aLayer.Editable Then
            If aLayer.Shapes.All.Count > 0 Then
                Set aSel = aLayer.Shapes.All
                Set shPowerClip = aLayer.CreateRectangle(sL, sT, sr, sB)
                shPowerClip.Outline.SetNoOutline
                aSel.AddToPowerClip shPowerClip, cdrFalse
            End If
        End If
    Next aLayer
    aPage.GuidesLayer.Editable = guideL
End Sub

Sub UnGroupAllPages()
    Dim p As Page
    For Each p In ActiveDocument.Pages
        p.Shapes.All.Ungroup
    Next p
End Sub

Sub MoveOnNewLayerOnAllPages()
    Dim p As Page
    Dim l As Layer
    Dim l2 As Layer
    Dim sr As ShapeRange
    
    For Each p In ActiveDocument.Pages
        p.Activate
        Set l = p.CreateLayer("Base Layer")
        Set sr = p.SelectableShapes.All
        sr.MoveToLayer l
        For Each l2 In p.Layers
            If l.Name <> l2.Name And l2.Index <> 0 Then
                l2.Delete
            End If
        Next l2
    Next p
End Sub

Sub TotalArea()
    Dim s As Shape
    Dim sum As Long
    For Each s In ActiveSelection.Shapes
        sum = sum + s.BoundingBox.height * s.BoundingBox.width
    Next s
    MsgBox sum
End Sub

