Attribute VB_Name = "bookletGuides"
Sub bookletGuides()
If (Documents.Count = 0) Then Exit Sub
    Application.ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    Dim guide As Shape
   
    'first page
    ActiveDocument.Pages(1).Activate
    ActiveDocument.Pages(1).SetSize 210, 296
    ActiveDocument.Pages(1).Orientation = cdrLandscape
    ActiveDocument.ActivePage.GuidesLayer.Editable = True
    ActiveDocument.ActivePage.GuidesLayer.Activate
    Set guide = ActiveDocument.ActivePage.GuidesLayer.CreateGuideAngle(0#, 0#, 90#)
    guide.Outline.SetProperties Color:=CreateRGBColor(0, 0, 255)
    Set guide = ActiveDocument.ActivePage.GuidesLayer.CreateGuideAngle(97#, 0#, 90#)
    guide.Outline.SetProperties Color:=CreateRGBColor(0, 0, 255)
    Set guide = ActiveDocument.ActivePage.GuidesLayer.CreateGuideAngle(196#, 0#, 90#)
    guide.Outline.SetProperties Color:=CreateRGBColor(0, 0, 255)
    Set guide = ActiveDocument.ActivePage.GuidesLayer.CreateGuideAngle(296#, 0#, 90#)
    guide.Outline.SetProperties Color:=CreateRGBColor(0, 0, 255)
    ActivePage.GuidesLayer.Editable = False
    ActivePage.GuidesLayer.Visible = True
    
    'second page
    If ActiveDocument.Pages.Count < 2 Then
        ActiveDocument.InsertPagesEx 1, False, ActivePage.Index, 296#, 210#
    End If
    ActiveDocument.Pages(2).Activate
    ActiveDocument.Pages(2).SetSize 210, 296
    ActiveDocument.Pages(2).Orientation = cdrLandscape
    ActiveDocument.ActivePage.GuidesLayer.Editable = True
    ActiveDocument.ActivePage.GuidesLayer.Activate
    Set guide = ActiveDocument.ActivePage.GuidesLayer.CreateGuideAngle(0#, 0#, 90#)
    guide.Outline.SetProperties Color:=CreateRGBColor(0, 0, 255)
    Set guide = ActiveDocument.ActivePage.GuidesLayer.CreateGuideAngle(296#, 0#, 90#)
    guide.Outline.SetProperties Color:=CreateRGBColor(0, 0, 255)
    Set guide = ActiveDocument.ActivePage.GuidesLayer.CreateGuideAngle(199#, 0#, 90#)
    guide.Outline.SetProperties Color:=CreateRGBColor(0, 0, 255)
    Set guide = ActiveDocument.ActivePage.GuidesLayer.CreateGuideAngle(100#, 0#, 90#)
    guide.Outline.SetProperties Color:=CreateRGBColor(0, 0, 255)

    ActiveDocument.ActivePage.GuidesLayer.Editable = False
    ActivePage.GuidesLayer.Visible = True
    ActivePage.Layers(2).Activate
    
    ActiveDocument.Pages(1).Activate
    ActiveDocument.ActivePage.Layers(2).Activate
End Sub


    
