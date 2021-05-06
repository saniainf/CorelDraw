Attribute VB_Name = "Utilities"
Sub ClearTransparency()
    Dim sr As ShapeRange
    Dim s As Shape
    Dim i As Integer
    Set sr = ActiveSelection.Shapes.All
    Application.Optimization = True
    For i = 1 To sr.Count
        If sr(i).Transparency.Type <> cdrNoTransparency Then
            sr(i).Transparency.ApplyNoTransparency
        End If
    Next i
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Sub ConvertToRGB300()
    ActiveSelection.ConvertToBitmapEx cdrRGBColorImage, False, False, 300, cdrNormalAntiAliasing, True, True, 95
End Sub

Sub RepeatToAllShape()
    Dim s As Shape
    
    For Each s In ActiveSelection.Shapes.All
        s.CreateSelection
        ActiveDocument.Repeat
    Next s
    
End Sub

Sub UnlockContentPowerClip()
    Dim pwc As PowerClip
    Dim s As Shape
    
    For Each s In ActiveSelection.Shapes.All
        Set pwc = Nothing
        Set pwc = s.PowerClip
            If Not pwc Is Nothing Then
                s.PowerClip.ContentsLocked = False
            End If
    Next s
        
End Sub

Sub LockContentPowerClip()
    Dim pwc As PowerClip
    Dim s As Shape
    
    For Each s In ActiveSelection.Shapes.All
        Set pwc = Nothing
        Set pwc = s.PowerClip
            If Not pwc Is Nothing Then
                s.PowerClip.ContentsLocked = True
            End If
    Next s
        
End Sub

Sub ExtractPowerClip()
    Dim aPage As Page
    Dim aDoc As Document
    Dim s As Shape
    Dim pwc As PowerClip
    
    For Each aDoc In Application.Documents
    aDoc.Activate
        For Each aPage In aDoc.Pages
            For Each s In aPage.SelectableShapes
                Set pwc = Nothing
                Set pwc = s.PowerClip
                If Not pwc Is Nothing Then
                    s.CreateSelection
                    pwc.ExtractShapes
                    s.Delete
                End If
            Next s
        Next aPage
    Next aDoc
End Sub

Sub RenamePage()
    Dim aPage As Page
    Dim l As Layer
    Dim label As String
    Dim labelnew As String
    
    For Each aPage In ActiveDocument.Pages
        label = aPage.Name
        labelnew = Replace(label, "/", "-")
        aPage.Name = labelnew
    Next aPage
End Sub

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

Sub InsertPageNumber()
    Dim aPage As Page
    Dim sNumber As Shape
    Dim x As Double
    Dim y As Double
    x = ActiveShape.PositionX
    y = ActiveShape.PositionY
    For Each aPage In ActiveDocument.Pages
        aPage.Activate
        Set sNumber = aPage.ActiveLayer.CreateArtisticText(x, y, aPage.Index, cdrLanguageNone, cdrCharSetMixed, "Arial", 9, cdrFalse, cdrFalse, cdrNoFontLine, cdrCenterAlignment)
    Next aPage
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

Sub CloseAllDocWithSaveOptions()
    Dim opt As New StructSaveAsOptions
    opt.EmbedICCProfile = False
    opt.EmbedVBAProject = False
    opt.Filter = cdrCDR
    opt.IncludeCMXData = False
    opt.Overwrite = True
    opt.Range = cdrAllPages
    opt.ThumbnailSize = cdr10KColorThumbnail
    opt.Version = cdrVersion17
    
    Dim doc As Document
    For Each doc In Application.Documents
        doc.SaveAs doc.fullFileName, opt
        doc.Close
    Next doc
End Sub

Sub CloseAllDocWithoutSave()
        Dim doc As Document
        For Each doc In Application.Documents
            doc.Dirty = False
            doc.Close
        Next doc
End Sub

Sub RenameFileAndSave()
        Dim doc As Document
        Dim oldName As String
        Dim newName As String
        Dim Path As String
        Dim colorCount As String
        Dim docPageSize As String
        For Each doc In Application.Documents
            doc.Activate
            doc.Unit = cdrMillimeter
            Path = doc.filePath
            oldName = doc.fileName
            oldName = Left(oldName, (Len(oldName) - 4))
            'size -1 mm
            docPageSize = doc.Pages.First.BoundingBox.width - 2 & "x" & doc.Pages.First.BoundingBox.height - 2
            'color
            If doc.Pages.Count > 1 Then
                colorCount = "4+4"
            Else
                colorCount = "4+0"
            End If
            
            newName = Path & oldName & "_" & colorCount & "_" & docPageSize & ".cdr"
            doc.SaveAs newName
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
                aPage.SizeWidth = aPage.SizeWidth - 4
                aPage.SizeHeight = aPage.SizeHeight - 4
'                aPage.SizeWidth = 214
'                aPage.SizeHeight = 102
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

            For Each aPage In doc.Pages 'paste on all page
                aPage.Activate
                Set s = aPage.ActiveLayer.Paste
'                s.OrderToBack 'order
            Next aPage
            
'            Set aPage = doc.Pages(1) 'paste only this page
'            aPage.Activate
'            Set s = aPage.ActiveLayer.Paste
'            s.OrderToBack 'order

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

Sub GroupAllPages()
    Dim p As Page
    For Each p In ActiveDocument.Pages
        p.Activate
        If p.SelectableShapes.Count > 0 Then
            p.SelectableShapes.All.Group
        End If
    Next p
End Sub

Sub UnGroupAllPages()
    Dim p As Page
    For Each p In ActiveDocument.Pages
        p.Shapes.All.UngroupAll
    Next p
End Sub

Sub UnGroupAllDoc()
    Dim d As Document
    Dim p As Page
    For Each d In Documents
        d.Activate
        For Each p In d.Pages
            p.Activate
            p.Shapes.All.UngroupAll
        Next p
    Next d
End Sub

Sub ConvertToCurves()
    Dim d As Document
    Dim p As Page
    For Each d In Documents
        d.Activate
        For Each p In d.Pages
            p.Activate
            p.Shapes.All.ConvertToCurves
        Next p
    Next d
End Sub

Sub MoveOnNewLayerOnAllPages()
    Dim p As Page
    Dim l As Layer
    Dim l2 As Layer
    Dim sr As ShapeRange
    
    For Each p In ActiveDocument.Pages
        p.Activate
        Set l = p.CreateLayer("Layer1")
        Set sr = p.SelectableShapes.All
        sr.MoveToLayer l
        sr.Group
        For Each l2 In p.Layers
            If l.Name <> l2.Name And l2.Index <> 0 Then
                l2.Delete
            End If
        Next l2
    Next p
End Sub

Sub MoveOnNewDocument()
    Dim cDoc As Document
    Dim newDoc As Document
    Dim sr As ShapeRange
    
    Set cDoc = ActiveDocument
    Set newDoc = CreateDocument()
    newDoc.InsertPages cDoc.Pages.Count - 1, False, 1
    iPage = 1
    
    For i = 1 To cDoc.Pages.Count
        cDoc.Pages(i).Activate
        cDoc.Pages(i).SelectableShapes.All.Copy
        newDoc.Pages(i).Activate
        newDoc.Pages(i).Name = cDoc.Pages(i).Name
        newDoc.Pages(i).Orientation = cDoc.Pages(i).Orientation
        newDoc.ActivePage.SetSize cDoc.ActivePage.SizeWidth, cDoc.ActivePage.SizeHeight
        newDoc.Pages(i).ActiveLayer.Paste
    Next i
End Sub

Sub TotalArea()
    Dim s As Shape
    Dim sum As Long
    For Each s In ActiveSelection.Shapes
        sum = sum + s.BoundingBox.height * s.BoundingBox.width
    Next s
    MsgBox sum
End Sub

Sub MainExportToPDF(doc As Document)
    With doc.PDFSettings
        .PublishRange = pdfWholeDocument
'        .PageRange = "1"
        .Author = ""
        .Subject = ""
        .Keywords = ""
        .BitmapCompression = 1 ' CdrPDFVBA.pdfLZW
'        .JPEGQualityFactor = 2
        .TextAsCurves = True
'        .EmbedFonts = True
'        .EmbedBaseFonts = True
'        .TrueTypeToType1 = True
'        .SubsetFonts = True
'        .SubsetPct = 80
        .CompressText = True
        .Encoding = 1 ' CdrPDFVBA.pdfBinary
        .DownsampleColor = False
        .DownsampleGray = False
        .DownsampleMono = False
'        .ColorResolution = 300
'        .MonoResolution = 1200
'        .GrayResolution = 300
        .Hyperlinks = False
        .Bookmarks = False
        .Thumbnails = False
        .Startup = 0 ' CdrPDFVBA.pdfPageOnly
        .ComplexFillsAsBitmaps = False
        .Overprints = False
        .Halftones = False
        .MaintainOPILinks = False
        .FountainSteps = 256
        .EPSAs = 0 ' CdrPDFVBA.pdfPostscript
        .pdfVersion = 9 ' CdrPDFVBA.pdfVersion17_Acrobat9
        .IncludeBleed = False
'        .Bleed = 31750
        .Linearize = False
        .CropMarks = False
        .RegistrationMarks = False
        .DensitometerScales = False
        .FileInformation = False
        .ColorMode = pdfNative
'        .UseColorProfile = True
'        .ColorProfile = 1 ' CdrPDFVBA.pdfSeparationProfile
        .EmbedFilename = ""
        .EmbedFile = False
        .JP2QualityFactor = 2
        .TextExportMode = 0 ' CdrPDFVBA.pdfTextAsUnicode
        .PrintPermissions = 0 ' CdrPDFVBA.pdfPrintPermissionNone
        .EditPermissions = 0 ' CdrPDFVBA.pdfEditPermissionNone
        .ContentCopyingAllowed = False
'        .OpenPassword = ""
'        .PermissionPassword = ""
        .EncryptType = 3 ' CdrPDFVBA.pdfEncryptTypeAES256
        .OutputSpotColorsAs = 0 ' CdrPDFVBA.pdfSpotAsSpot
'        .OverprintBlackLimit = 95
    End With

    Dim fileName As String
    Dim filePath As String
    Dim fullFileName As String
    
    fileName = doc.fileName
    filePath = doc.filePath
    fileName = Left(fileName, (Len(fileName) - 4))
    fullFileName = filePath & fileName & ".pdf"
    
    doc.PublishToPDF fullFileName
End Sub

Sub ExportAllToPDF()
    Dim doc As Document
    For Each doc In Documents
        MainExportToPDF doc
    Next doc
End Sub

Sub ConvertToGray(sr As ShapeRange)
    Dim s As Shape
    
    For Each s In sr
        If s.Type = cdrGroupShape Then
            ConvertToGray s.Shapes.All
        End If
        If s.Type = cdrBitmapShape Then
            s.Bitmap.ConvertTo cdrGrayscaleImage
            s.Bitmap.ApplyBitmapEffect "Automatically adjust color and contrast", "AutoEqualizeEffect "
            s.Bitmap.ConvertTo cdrCMYKColorImage
        End If
    Next s
End Sub

Sub ConvertToGraySelectedShapes()
    ConvertToGray ActiveSelection.Shapes.All
End Sub

Sub ConvertToCMYKProfile(sr As ShapeRange)
    Dim s As Shape
    
    For Each s In sr
        If s.Type = cdrGroupShape Then
            ConvertToCMYKProfile s.Shapes.All
        End If
        If s.Type = cdrBitmapShape Then
            s.Bitmap.ConvertTo cdrLABImage
            s.Bitmap.ConvertTo cdrCMYKColorImage
        End If
    Next s
End Sub

Sub ConvertToCMYKProfileAllPages()
    Dim p As Page
    
    For Each p In ActiveDocument.Pages
        p.Activate
        ConvertToCMYKProfile p.SelectableShapes.All
    Next p
End Sub

