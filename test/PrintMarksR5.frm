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
Dim cColor As New Collection, cColorOnly As New Collection
Dim cyanColor As New Color, magentaColor As New Color, yellowColor As New Color, blackColor As New Color
Dim whiteColor As New Color
Dim cBlack40 As New Color, cGrayBalance As New Color
Dim tint80 As String, tint40 As String, grayBalance As String, black40 As String
Dim prevSelected As Integer
Dim clr As New Color
Dim i As Integer, e As Integer, a As Integer
Dim obj As Variant
Dim typeStr As Boolean
Dim str As String

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
        If TypeName(cColor.Item(prevSelected + 1)) = "String" Then
            str = cColor.Item(prevSelected + 1)
            typeStr = True
        End If
        If TypeName(cColor.Item(prevSelected + 1)) = "IDrawColor" Then
            Set clr = cColor.Item(prevSelected + 1)
            typeStr = False
        End If

        cColor.Remove prevSelected + 1
        If typeStr Then
            cColor.Add str, , , prevSelected + 1
        Else
            cColor.Add clr, , , prevSelected + 1
        End If
        lbColorList.ListIndex = sbColorList.Value
        colorListUpdate
    End If

    If prevSelected > sbColorList.Value Then
        If TypeName(cColor.Item(prevSelected + 1)) = "String" Then
            str = cColor.Item(prevSelected + 1)
            typeStr = True
        End If
        If TypeName(cColor.Item(prevSelected + 1)) = "IDrawColor" Then
            Set clr = cColor.Item(prevSelected + 1)
            typeStr = False
        End If

        cColor.Remove prevSelected + 1
        If typeStr Then
            cColor.Add str, , prevSelected
        Else
            cColor.Add clr, , prevSelected
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
    If clr.UserAssignEx Then
        If clr.Name = "unnamed color" Then
            MsgBox "Unnamed Color", vbCritical, "Error"
            Exit Sub
        End If
        'if not selected
        If lbColorList.ListIndex = -1 Then lbColorList.ListIndex = lbColorList.ListCount - 1
        cColor.Add clr, , lbColorList.ListIndex + 1
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

End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Sub colorListUpdate()
    'save selected item
    i = lbColorList.ListIndex
    
    lbColorList.Clear
    Set cColorOnly = New Collection
    
    For Each obj In cColor
        Select Case TypeName(obj)
            Case "IDrawColor"
                lbColorList.AddItem obj.Name
                cColorOnly.Add obj
            Case "String"
                lbColorList.AddItem parserStringToColorList(obj)
        End Select
    Next obj
    
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

Sub fillLabel()
    lblBarCount.Caption = "Bar count: " & cColor.Count
    i = 0
    e = 0
    For Each clr In cColorOnly
        If clr.Type = cdrColorCMYK Then
            i = i + 1
        ElseIf clr.Type = cdrColorSpot Or clr.Type = cdrColorPantone Then
            e = e + 1
        End If
    Next clr
    lblCmykCount.Caption = "CMYK count: " & i
    lblSpotCount.Caption = "Spot count: " & e
End Sub
