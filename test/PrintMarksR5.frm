VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PrintMarksR5 
   Caption         =   "Print Marks Ryobi 5xx"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5415
   OleObjectBlob   =   "PrintMarksR5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PrintMarksR5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cColor As New Collection, cColorOnly As New Collection, cColorBar As New Collection, cColorSign As New Collection
Dim cyanColor As New Color, magentaColor As New Color, yellowColor As New Color, blackColor As New Color
Dim whiteColor As New Color
Dim cBlack40 As New Color, cGrayBalance As New Color
Dim tint80 As String, tint40 As String, grayBalance As String, black40 As String
Dim x As Double, spaceWidth As Double, barWidth As Double, startPos As Double
Dim sBar As Shape, sText As Shape, sSign As Shape
Dim sectionCount As Integer, barInSection As Integer
Dim prevSelected As Integer
Dim oClr As New Color, saveClr As New Color, pickClr As New Color, tintClr As New Color
Dim i As Integer, e As Integer, a As Integer
Dim objCColorForList As Variant, objCColorForBar As Variant
Dim typeStr As Boolean
Dim str As String, saveStr As String
Dim icClr As Integer, icClrOnly As Integer
Dim srBar As New ShapeRange
Dim rPoint As cdrReferencePoint
Dim listHeight As Double

Private Sub UserForm_Initialize()
    Application.ActiveDocument.Unit = cdrMillimeter
    
    cyanColor.CMYKAssign 100, 0, 0, 0
    magentaColor.CMYKAssign 0, 100, 0, 0
    yellowColor.CMYKAssign 0, 0, 100, 0
    blackColor.CMYKAssign 0, 0, 0, 100
    whiteColor.CMYKAssign 0, 0, 0, 0
    cBlack40.CMYKAssign 0, 0, 0, 40
    cGrayBalance.CMYKAssign 38, 26, 26, 0
    grayBalance = "grayBalance"
    black40 = "black40"
    tint80 = "tint80"
    tint40 = "tint40"
    
    sectionCount = 16
    barInSection = 8
    spaceWidth = 0.3
    barWidth = 3.9
    startPos = 0#
    listHeight = lbColorList.Height
    
    cColor.Add cyanColor
    cColor.Add magentaColor
    cColor.Add yellowColor
    cColor.Add blackColor
    cColor.Add grayBalance
    cColor.Add black40
    cColor.Add tint80
    cColor.Add tint40
    
    colorListUpdate
End Sub

Private Sub lbColorList_Change()
    If lbColorList.ListCount <= 0 Or lbColorList.ListIndex < 0 Then Exit Sub
    prevSelected = lbColorList.ListIndex
End Sub

Private Sub lbColorList_Click()
    If lbColorList.ListCount <= 0 Or lbColorList.ListIndex < 0 Then Exit Sub
    sbColorList.Value = lbColorList.ListIndex
End Sub

Private Sub sbColorList_Change()
    If lbColorList.ListCount <= 0 Or lbColorList.ListIndex < 0 Then Exit Sub
   
    If prevSelected < sbColorList.Value Then
        'save color or str
        If TypeName(cColor.Item(prevSelected + 1)) = "String" Then
            saveStr = cColor.Item(prevSelected + 1)
            typeStr = True
        End If
        If TypeName(cColor.Item(prevSelected + 1)) = "IDrawColor" Then
            Set saveClr = cColor.Item(prevSelected + 1)
            typeStr = False
        End If
        'remove item
        cColor.Remove prevSelected + 1
        'add item in new position
        If typeStr Then
            cColor.Add saveStr, , , prevSelected + 1
        Else
            cColor.Add saveClr, , , prevSelected + 1
        End If
        lbColorList.ListIndex = sbColorList.Value
        colorListUpdate
    End If

    If prevSelected > sbColorList.Value Then
        If TypeName(cColor.Item(prevSelected + 1)) = "String" Then
            saveStr = cColor.Item(prevSelected + 1)
            typeStr = True
        End If
        If TypeName(cColor.Item(prevSelected + 1)) = "IDrawColor" Then
            Set saveClr = cColor.Item(prevSelected + 1)
            typeStr = False
        End If

        cColor.Remove prevSelected + 1
        If typeStr Then
            cColor.Add saveStr, , prevSelected
        Else
            cColor.Add saveClr, , prevSelected
        End If
        lbColorList.ListIndex = sbColorList.Value
        colorListUpdate
    End If
End Sub
'Private Sub sbColorList_Change()
'    If lbColorList.ListCount <= 0 Or lbColorList.ListIndex < 0 Then Exit Sub
'    Dim vObj As Variant
'
'    If prevSelected < sbColorList.Value Then
'        vObj = cColor.Item(prevSelected + 1)
'        cColor.Remove prevSelected + 1
'        cColor.Add vObj, , , prevSelected + 1
'        lbColorList.ListIndex = sbColorList.Value
'        colorListUpdate
'    End If
'
'    If prevSelected > sbColorList.Value Then
'        vObj = cColor.Item(prevSelected + 1)
'        cColor.Remove prevSelected + 1
'        cColor.Add vObj, , prevSelected
'        lbColorList.ListIndex = sbColorList.Value
'        colorListUpdate
'    End If
'End Sub

Private Sub btnAddColor_Click()
    Set pickClr = New Color
    If pickClr.UserAssignEx Then
        If pickClr.Name = "unnamed color" Then
            MsgBox "Unnamed Color", vbCritical, "Error"
            Exit Sub
        End If
        'if not selected
        If lbColorList.ListIndex = -1 Then lbColorList.ListIndex = lbColorList.ListCount - 1
        cColor.Add pickClr, , lbColorList.ListIndex + 1
        colorListUpdate
    End If
End Sub

Private Sub btnDeleteColor_Click()
    If lbColorList.ListCount = 1 Then Exit Sub
    'if not selected
    If lbColorList.ListIndex = -1 Then lbColorList.ListIndex = lbColorList.ListCount - 1
    cColor.Remove lbColorList.ListIndex + 1
    colorListUpdate
End Sub

Private Sub btnAddCyan_Click()
    cColor.Add cyanColor, , lbColorList.ListIndex + 1
    colorListUpdate
End Sub

Private Sub btnAddMagenta_Click()
    cColor.Add magentaColor, , lbColorList.ListIndex + 1
    colorListUpdate
End Sub

Private Sub btnAddYellow_Click()
    cColor.Add yellowColor, , lbColorList.ListIndex + 1
    colorListUpdate
End Sub

Private Sub btnAddBlack_Click()
    cColor.Add blackColor, , lbColorList.ListIndex + 1
    colorListUpdate
End Sub

Private Sub btnAdd40_Click()
    cColor.Add tint40, , lbColorList.ListIndex + 1
    colorListUpdate
End Sub

Private Sub btnAdd80_Click()
    cColor.Add tint80, , lbColorList.ListIndex + 1
    colorListUpdate
End Sub

Private Sub btnAddGrayBalance_Click()
    cColor.Add black40, , lbColorList.ListIndex + 1
    cColor.Add grayBalance, , lbColorList.ListIndex + 1
    colorListUpdate
End Sub

Private Sub btnCreateMarks_Click()
If (Documents.Count = 0) Then
    Unload Me
    Exit Sub
End If
    ActiveDocument.BeginCommandGroup "Ñreate Print Marks"
    rPoint = ActiveDocument.ReferencePoint
    ActiveDocument.ReferencePoint = cdrTopLeft
    Application.Optimization = True
    Set cColorBar = New Collection
    Set cColorSign = New Collection
    
    If cColor.Count <= 8 Then createStandartBar
    If cColor.Count > 8 Then createExtendetBar
    
    'color bar
    signBar
    Set srBar = New ShapeRange
    ActiveDocument.ClearSelection
    For Each sBar In cColorBar
        srBar.Add sBar
    Next sBar
    Set sBar = srBar.Group
    
    'sign color
    signColor
    Set srBar = New ShapeRange
    ActiveDocument.ClearSelection
    For Each sSign In cColorSign
        srBar.Add sSign
    Next sSign
    Set sSign = srBar.Group
    
    PrintMarksR5v2.PrintMarksR5v2 sBar, sSign
    
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
    ActiveDocument.ReferencePoint = rPoint
    ActiveDocument.EndCommandGroup
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Sub colorListUpdate()
    'save selected item
    i = lbColorList.ListIndex
    
    lbColorList.Clear
    Set cColorOnly = New Collection
    
    For Each objCColorForList In cColor
        Select Case TypeName(objCColorForList)
            Case "IDrawColor"
                lbColorList.AddItem objCColorForList.Name
                cColorOnly.Add objCColorForList
            Case "String"
                lbColorList.AddItem parserStringToColorList(objCColorForList)
        End Select
    Next objCColorForList
    
    'if delete item and reduce items count
    If i >= lbColorList.ListCount Then i = lbColorList.ListCount - 1
    'restore selected item
    lbColorList.ListIndex = i
    'if not selected
    If lbColorList.ListIndex = -1 Then lbColorList.ListIndex = lbColorList.ListCount - 1
    'label fill
    fillLabel
    'scroll color list min_max
    sbColorList.Min = 0
    sbColorList.Max = lbColorList.ListCount - 1
End Sub

Sub createExtendetBar()
    x = startPos
    icClr = 1
    icClrOnly = 1
    For i = 0 To sectionCount - 1
        'color bar
        For a = 1 To barInSection
            Set sBar = ActiveLayer.CreateRectangle(x, barWidth, x + barWidth, 0)
            cColorBar.Add sBar
            sBar.Outline.SetNoOutline
            If TypeName(cColor.Item(icClr)) = "IDrawColor" Then
                sBar.Fill.UniformColor = cColor.Item(icClr)
            Else
                sBar.Fill.UniformColor = parserStringToExtColorBar(cColor.Item(icClr))
            End If
            x = x + barWidth
            icClr = icClr + 1
            If icClrOnly > cColorOnly.Count Then icClrOnly = 1
            If icClr > cColor.Count Then icClr = 1
        Next a
        'white space
        Set sBar = ActiveLayer.CreateRectangle(x, barWidth, x + spaceWidth, 0)
        cColorBar.Add sBar
        sBar.Outline.SetNoOutline
        sBar.Fill.ApplyNoFill
        x = x + spaceWidth
    Next i
End Sub

Sub createStandartBar()
    x = startPos
    For i = 0 To sectionCount - 1
        'color bar
        For a = 1 To barInSection \ cColor.Count
            For Each objCColorForBar In cColor
                Set sBar = ActiveLayer.CreateRectangle(x, barWidth, x + barWidth, 0)
                cColorBar.Add sBar
                sBar.Outline.SetNoOutline
                If TypeName(objCColorForBar) = "IDrawColor" Then
                    sBar.Fill.UniformColor = objCColorForBar
                Else
                    sBar.Fill.UniformColor = parserStringToColorBar(objCColorForBar, i)
                End If
                x = x + barWidth
            Next objCColorForBar
        Next a
        'white bar
        For a = 0 To barInSection Mod cColor.Count - 1
            Set sBar = ActiveLayer.CreateRectangle(x, barWidth, x + barWidth, 0)
            cColorBar.Add sBar
            sBar.Outline.SetNoOutline
            sBar.Fill.ApplyNoFill
            x = x + barWidth
        Next a
        'white space
        Set sBar = ActiveLayer.CreateRectangle(x, barWidth, x + spaceWidth, 0)
        cColorBar.Add sBar
        sBar.Outline.SetNoOutline
        sBar.Fill.ApplyNoFill
        x = x + spaceWidth
    Next i
End Sub

Private Sub cbCustomSign_Click()
    If cbCustomSign.Value = True Then
        lbColorList.ListStyle = fmListStyleOption
        colorListUpdate
        lbColorList.Height = listHeight
    Else
        lbColorList.ListStyle = fmListStylePlain
        colorListUpdate
        lbColorList.Height = listHeight
    End If
End Sub

Private Sub signBar()
    If cbCustomSign.Value Then
        If lbColorList.ListIndex + 1 > barInSection Then
            x = startPos + barWidth * barInSection / 2
        Else
            x = startPos + barWidth * lbColorList.ListIndex
        End If
    Else
        x = startPos + barWidth * barInSection / 2
    End If
    For i = sectionCount - 1 To 0 Step -1
        Set sText = ActiveLayer.CreateArtisticText(x, barWidth, i + 1, , , "Arial", 4, cdrFalse, cdrFalse, cdrNoFontLine, cdrLeftAlignment)
        sText.ConvertToCurves
        sText.PositionY = sText.PositionY - sText.BoundingBox.Height
        sText.Fill.UniformColor = whiteColor
        sText.Outline.SetNoOutline
        cColorBar.Add sText
        x = x + barWidth * barInSection + spaceWidth
    Next i
End Sub

Private Sub signColor()
    x = startPos
    For Each oClr In cColorOnly
        Set sText = ActiveLayer.CreateArtisticText(x, 0, oClr.Name, , , "Arial", 6, cdrTrue, cdrFalse, cdrNoFontLine, cdrLeftAlignment)
        sText.ConvertToCurves
        sText.Fill.UniformColor = oClr
        sText.Outline.SetNoOutline
        cColorSign.Add sText
        x = x + sText.BoundingBox.Width + barWidth / 2
    Next oClr
End Sub

Public Function parserStringToColorList(pStr As Variant) As String
    Select Case pStr
        Case grayBalance
            parserStringToColorList = "Gray Balance"
        Case black40
            parserStringToColorList = "K: 40%"
        Case tint80
            parserStringToColorList = "Tint: 80%"
        Case tint40
            parserStringToColorList = "Tint: 40%"
    End Select
End Function

Public Function parserStringToExtColorBar(pStr As Variant) As Color
    Select Case pStr
        Case grayBalance
            Set parserStringToExtColorBar = cGrayBalance
        Case black40
            Set parserStringToExtColorBar = cBlack40
        Case tint80
            'get copy color from ColorOnly collection
            Set tintClr = cColorOnly.Item(icClrOnly).GetCopy
            'tint color differently for spot or cmyk
            If tintClr.Type = cdrColorCMYK Then
                tintClr.BlendWith whiteColor, 80
            ElseIf tintClr.Type = cdrColorSpot Or tintClr.Type = cdrColorPantone Then
                Set tintClr = CreateSpotColor(tintClr.PaletteIdentifier, tintClr.SpotColorID, 80)
            End If
            icClrOnly = icClrOnly + 1
            'return value
            Set parserStringToExtColorBar = tintClr
        Case tint40
            'get copy color from ColorOnly collection
            Set tintClr = cColorOnly.Item(icClrOnly).GetCopy
            'tint color differently for spot or cmyk
            If tintClr.Type = cdrColorCMYK Then
                tintClr.BlendWith whiteColor, 40
            ElseIf tintClr.Type = cdrColorSpot Or tintClr.Type = cdrColorPantone Then
                Set tintClr = CreateSpotColor(tintClr.PaletteIdentifier, tintClr.SpotColorID, 40)
            End If
            icClrOnly = icClrOnly + 1
            'return value
            Set parserStringToExtColorBar = tintClr
    End Select
End Function

Public Function parserStringToColorBar(pStr As Variant, nSection As Integer) As Color
    Select Case pStr
        Case grayBalance
            Set parserStringToColorBar = cGrayBalance
        Case black40
            Set parserStringToColorBar = cBlack40
        Case tint80
            'get copy color from ColorOnly collection
            Set tintClr = cColorOnly.Item((nSection Mod cColorOnly.Count) + 1).GetCopy
            'tint color differently for spot or cmyk
            If tintClr.Type = cdrColorCMYK Then
                tintClr.BlendWith whiteColor, 80
            ElseIf tintClr.Type = cdrColorSpot Or tintClr.Type = cdrColorPantone Then
                Set tintClr = CreateSpotColor(tintClr.PaletteIdentifier, tintClr.SpotColorID, 80)
            End If
            'return value
            Set parserStringToColorBar = tintClr
        Case tint40
            Set tintClr = cColorOnly.Item((nSection Mod cColorOnly.Count) + 1).GetCopy
            If tintClr.Type = cdrColorCMYK Then
                tintClr.BlendWith whiteColor, 40
            ElseIf tintClr.Type = cdrColorSpot Or tintClr.Type = cdrColorPantone Then
                Set tintClr = CreateSpotColor(tintClr.PaletteIdentifier, tintClr.SpotColorID, 40)
            End If
            Set parserStringToColorBar = tintClr
    End Select
End Function

Sub fillLabel()
    If cColor.Count > barInSection Then
        lblBarCount.ForeColor = &HFF&
    Else
        lblBarCount.ForeColor = &H80000012
    End If
    
    lblBarCount.Caption = "Bar count: " & cColor.Count
    i = 0
    e = 0
    For Each oClr In cColorOnly
        If oClr.Type = cdrColorCMYK Then
            i = i + 1
        ElseIf oClr.Type = cdrColorSpot Or oClr.Type = cdrColorPantone Then
            e = e + 1
        End If
    Next oClr
    lblCmykCount.Caption = "CMYK count: " & i
    lblSpotCount.Caption = "Spot count: " & e
End Sub
