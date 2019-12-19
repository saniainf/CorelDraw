Attribute VB_Name = "RecordedMacros"
Sub Macro1()
 Dim s As Shape
 Set s = ActiveLayer.CreateRectangle(1, 1, 4, 3)
 With s.Fill.ApplyFountainFill(CreateRGBColor(0, 0, 0), CreateRGBColor(255, 0, 0))
  .Colors.Add CreateCMYKColor(0, 50, 50, 0), 77
  .Colors.Add CreateRGBColor(0, 255, 255), 60
 End With
End Sub

Sub testMacro()
    Dim s1 As Shape
    For Each s1 In ActiveDocument.SelectableShapes.All
        'MsgBox s1.Fill.UniformColor.Tint
        's1.Fill.UniformColor.SpotAssignByName "fc6ccd74-10a8-4152-8901-a51facb47785", "PANTONE 7466 C", 38
    Next s1
End Sub

Sub testProp()
 Const MyMacroName As String = "MyTestMacro"
 With ActiveDocument
  MsgBox .Properties(MyMacroName, 1)
  MsgBox .Properties(MyMacroName, 2)
 MsgBox .Properties(MyMacroName, 3)
 End With
End Sub


Sub oleconv()
    Application.ActiveDocument.Unit = cdrMillimeter
    Dim s As Shape
    Dim pv As Shape
    Dim rect As Shape
    Dim rect1 As Shape
    Dim rect2 As Shape
    Dim tn As TreeNode
    For Each s In ActiveDocument.SelectableShapes.All
        If s.Type = cdrOLEObjectShape Then
'            Set rect1 = ActiveLayer.CreateRectangle(s.LeftX, s.TopY, s.RightX, s.BottomY)
'            Set rect2 = ActiveLayer.CreateRectangle(s.LeftX, s.TopY, s.RightX, s.BottomY)
'            rect1.Fill.UniformColor.CMYKAssign 0, 100, 0, 0
'            rect2.Fill.UniformColor.CMYKAssign 100, 100, 0, 0
'            rect1.AddToSelection
'            rect2.AddToSelection
'            Set rect = ActiveDocument.Selection.Group
'            rect.TreeNode.MoveAfter s.TreeNode
            Set tn = s.TreeNode
            s.Copy
            ActiveLayer.PasteSpecial ("Metafile")
            Set pv = ActiveSelection.Shapes.First
            pv.TreeNode.MoveAfter tn
        End If
    Next s
End Sub

Sub deleteMagenta()
    Dim doc As Document
    Dim aPage As Page
    Dim s As Shape
    Dim c As Color
    
    Set c = New Color
    c.CMYKAssign 0, 100, 0, 0
   
    For Each doc In Application.Documents
        doc.Activate
        For Each aPage In doc.Pages
            For Each s In ActiveDocument.SelectableShapes.All
                If s.Fill.Type = cdrUniformFill Then
                    If s.Fill.UniformColor.IsSame(c) Then
                        s.Fill.ApplyNoFill
                    End If
                End If
            Next s
        Next aPage
    Next doc
End Sub

Sub offLayer()
    Dim doc As Document
    Dim aPage As Page
    Dim l As Layer
  
    For Each doc In Application.Documents
        doc.Activate
        For Each aPage In doc.Pages
            For Each l In aPage.AllLayers
                If l.Name = "LAK" Then
                    l.Visible = False
                    l.Printable = False
                End If
            Next l
        Next aPage
    Next doc
End Sub

Sub signCityName()
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

Sub SaveAllDoc()
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
                aPage.SizeHeight = 72
                aPage.SizeWidth = 102
            Next aPage
        Next doc
    
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub


Sub PlaceAllToPowerClip(aPage As Page)
    Dim aSel As ShapeRange
    Dim shPowerClip As Shape
    Dim sL As Integer
    Dim sT As Integer
    Dim sR As Integer
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
    sR = aPage.BoundingBox.Right
    sB = aPage.BoundingBox.Bottom
    For Each aLayer In aPage.Layers
        If aLayer.Editable Then
            If aLayer.Shapes.All.Count > 0 Then
                Set aSel = aLayer.Shapes.All
                Set shPowerClip = aLayer.CreateRectangle(sL, sT, sR, sB)
                shPowerClip.Outline.SetNoOutline
                aSel.AddToPowerClip shPowerClip, cdrFalse
            End If
        End If
    Next aLayer
    aPage.GuidesLayer.Editable = guideL
End Sub

Sub Testss()
 Dim Seps As SeparationPlates
 Dim p As SeparationPlate
 Dim i As Integer
 i = 0
    Set Seps = ActiveDocument.PrintSettings.Separations.Plates
    For Each p In Seps
        If p.Enabled Then
            i = i + 1
        End If
    Next p
    MsgBox "There are " & i & " plates."
End Sub

Sub Test111()
    Application.Optimization = True
    
        Dim doc As Document
        Dim aPage As Page
        Dim s As Shape
        For Each doc In Application.Documents
            doc.Activate
            doc.Unit = cdrMillimeter
            For Each aPage In doc.Pages
                Set s = aPage.Shapes.All.FirstShape
                s.Fill.ApplyNoFill
            Next aPage
        Next doc
    
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Sub TestText()
    Dim d As Document
    Dim s As Shape
    Dim t As Text
    Dim tr As TextRange
    Dim char As TextRange
    Dim textch As TextCharacters
    Dim i As Integer
                
    Set d = ActiveDocument
    Set s = d.ActiveLayer.CreateArtisticText(2, 2, "This isaa")
    Set textch = s.Text.Story.Characters
    i = 1
        
    For Each char In textch
        i = i + 10
        char.Fill.UniformColor.CMYKAssign 0, 0, 0, i
    Next char
End Sub
