Attribute VB_Name = "RecordedMacros"

Sub TemporaryMacro()
    ' Recorded 03.12.2015
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    Windows.FindWindow("Untitled-1").ActiveView.SetViewPoint 88.066, 249.8513, 63
    Windows.FindWindow("Untitled-1").ActiveView.SetViewPoint 103.7292, 131.7412, 31
    Windows.FindWindow("Untitled-1").ActiveView.ToFitArea 30.0694, 426.3813, 610.8829, -184.0654
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Redo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    Windows.FindWindow("Untitled-1").ActiveView.ToFitArea -26.811, 363.4593, 629.5235, -59.9598
    Windows.FindWindow("Untitled-1").ActiveView.ToFitAllObjects
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    Windows.FindWindow("Untitled-1").ActiveView.SetViewPoint 496.5308, 320.4399, 100
    Windows.FindWindow("Untitled-1").ActiveView.SetViewPoint 496.5307, 319.6463, 200
    Windows.FindWindow("Untitled-1").ActiveView.SetViewPoint 500.1026, 315.1483, 400
    ActiveDocument.Undo
    Windows.FindWindow("Untitled-1").ActiveView.ToFitAllObjects
    Windows.FindWindow("Untitled-1").ActiveView.SetViewPoint -14.7438, 115.6088, 38
    Windows.FindWindow("Untitled-1").ActiveView.ToFitArea -75.4217, 410.5312, 733.8508, -64.3077
    ActiveDocument.Undo
    Windows.FindWindow("Untitled-1").ActiveView.ToFitArea -56.1188, 356.7198, 259.1786, 175.0047
    Windows.FindWindow("Untitled-1").ActiveView.ToFitAllObjects
    ActiveDocument.Undo
    Windows.FindWindow("Untitled-1").ActiveView.ToFitArea 402.4044, 143.961, 524.2811, -26.0249
    ActiveDocument.Undo
    ActiveDocument.Undo
    Windows.FindWindow("Untitled-1").ActiveView.SetViewPoint 474.0879, 29.0559, 63
    Windows.FindWindow("Untitled-1").ActiveView.ToFitAllObjects
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    ActiveDocument.Undo
    Windows.FindWindow("Untitled-1").ActiveView.ToFitArea 420.0991, 117.4622, 523.0051, -37.4927
    Windows.FindWindow("Untitled-1").ActiveView.ToFitAllObjects
    Windows.FindWindow("Untitled-1").ActiveView.SetViewPoint 256.1553, 202.3962, 67
    Dim SaveOptions As StructSaveAsOptions
    Set SaveOptions = CreateStructSaveAsOptions
    With SaveOptions
        .EmbedVBAProject = True
        .Filter = cdrCDR
        .IncludeCMXData = False
        .Range = cdrAllPages
        .EmbedICCProfile = True
        .Version = cdrVersion15
    End With
    ActiveDocument.SaveAs "D:\work\12_Декабрь_2015\Фабрика Кухня\#4981.cdr", SaveOptions
    With ActiveDocument.PDFSettings
        .PublishRange = 0 ' CdrPDFVBA.pdfWholeDocument
        .PageRange = ""
        .Author = "Александр Бородич"
        .Subject = ""
        .Keywords = ""
        .BitmapCompression = 1 ' CdrPDFVBA.pdfLZW
        .JPEGQualityFactor = 2
        .TextAsCurves = True
        .EmbedFonts = True
        .EmbedBaseFonts = True
        .TrueTypeToType1 = True
        .SubsetFonts = True
        .SubsetPct = 80
        .CompressText = True
        .Encoding = 1 ' CdrPDFVBA.pdfBinary
        .DownsampleColor = False
        .DownsampleGray = False
        .DownsampleMono = False
        .ColorResolution = 300
        .MonoResolution = 1200
        .GrayResolution = 300
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
        .pdfVersion = 6 ' CdrPDFVBA.pdfVersion15
        .IncludeBleed = False
        .Bleed = 31750
        .Linearize = False
        .CropMarks = False
        .RegistrationMarks = False
        .DensitometerScales = False
        .FileInformation = False
        .ColorMode = 3 ' CdrPDFVBA.pdfNative
        .UseColorProfile = True
        .ColorProfile = 1 ' CdrPDFVBA.pdfSeparationProfile
        .EmbedFilename = ""
        .EmbedFile = False
        .JP2QualityFactor = 2
        .TextExportMode = 0 ' CdrPDFVBA.pdfTextAsUnicode
        .PrintPermissions = 0 ' CdrPDFVBA.pdfPrintPermissionNone
        .EditPermissions = 0 ' CdrPDFVBA.pdfEditPermissionNone
        .ContentCopyingAllowed = False
        .OpenPassword = ""
        .PermissionPassword = ""
        .EncryptType = 1 ' CdrPDFVBA.pdfEncryptTypeStandard
        .OutputSpotColorsAs = 0 ' CdrPDFVBA.pdfSpotAsSpot
        .OverprintBlackLimit = 95
    End With
    ActiveDocument.PublishToPDF "D:\PDF\out\#4981.pdf"
    ActiveDocument.Close
    AppWindow.WindowState = cdrWindowMinimized
    Windows.CloseAll
End Sub
