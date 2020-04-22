VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportAll 
   Caption         =   "Export All v2.04"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   OleObjectBlob   =   "frmExportAll.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExportAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'v 1.1
'correct export many doc
'v 1.2
'shapes -> SelectableShapes
'v 2.03
'correct file name
'v 2.04
'double numbering filename

Private Sub cbExecute_Click()
    Application.Optimization = True
    
    Dim doc As Document
    
    If (chbAllFiles) Then
        For Each doc In Documents
            doc.Activate
            exportJpeg doc
        Next doc
    Else
        Set doc = ActiveDocument
        exportJpeg doc
    End If
    
    ActiveDocument.Pages(1).Activate
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
    Unload Me
End Sub

Private Sub exportJpeg(doc As Document)
    Dim expFilter As ExportFilter
    Dim resolution As Integer
    Dim fileName As String
    Dim filePath As String
    Dim fullFileName As String
    Dim pageName As String
    Dim fileCount As String
    Dim colorSpace As Integer
    Dim colorSpaceField As String
    Dim aPage As Page
    Dim iPage As Integer
    Dim expArea As Rect
    
    fileName = doc.fileName
    filePath = doc.filePath
    fileName = Left(fileName, (Len(fileName) - 4))
    
    iPage = 0
    doc.MasterPage.GuidesLayer.Editable = False
    
    colorSpaceField = cboColorSpace.Text
    resolution = cboResolution.Text
    If (colorSpaceField = "Grayscale (8-bit)") Then colorSpace = 2
    If (colorSpaceField = "RGB Color (24-bit)") Then colorSpace = 4
    If (colorSpaceField = "CMYK color (24-bit)") Then colorSpace = 5
    
    For Each aPage In doc.Pages
        aPage.Activate
        aPage.GuidesLayer.Editable = False
        iPage = iPage + 1
        pageName = Replace_symbols(aPage.Name)
        fullFileName = filePath & fileName & "_" & Format(iPage, "0#") & "_" & pageName & ".jpg"
        aPage.SelectableShapes.All.CreateSelection
        If (aPage.SelectableShapes.Count > 0) Then
            Set expArea = aPage.SelectableShapes.All.BoundingBox.GetCopy
            If cbPageBox.Value Then
                Set expArea = aPage.BoundingBox.GetCopy
            End If
            Set expFilter = doc.ExportBitmap(fullFileName, cdrJPEG, cdrCurrentPage, colorSpace, 0, 0, resolution, resolution, cdrNormalAntiAliasing, False, False, chbProfile.Value, False, cdrCompressionNone, , expArea)
            With expFilter
                .Progressive = False
                .Optimized = False
                .SubFormat = 0
                .Compression = 10
                .Smoothing = 10
                .Finish
            End With
        End If
    Next aPage
End Sub

Private Sub cbCancel_Click()
    Unload Me
End Sub

Private Sub cbPageBox_Click()

End Sub

Private Sub UserForm_Initialize()
    cboResolution.AddItem "72"
    cboResolution.AddItem "150"
    cboResolution.AddItem "300"
    cboResolution.AddItem "1200"
    cboResolution.AddItem "2400"
    
    cboColorSpace.AddItem "Grayscale (8-bit)"
    cboColorSpace.AddItem "RGB Color (24-bit)"
    cboColorSpace.AddItem "CMYK color (24-bit)"
End Sub

Function Replace_symbols(ByVal txt As String) As String
    sT$ = "<>?.,%*!;':~!@/\#$%^&*=|`"""
    For i% = 1 To Len(sT$)
        txt = Replace(txt, Mid(sT$, i, 1), "_")
    Next
    Replace_symbols = txt
End Function

