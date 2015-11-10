VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrintSign 
   Caption         =   "Print Sign v1.0"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5625
   OleObjectBlob   =   "frmPrintSign.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrintSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public retval As Long
Public signX As Double, signY As Double, modShift As Long
Public sChar As String
Public customPlace As Boolean
Public offsetBottom As Integer

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnMake_Click()
    Application.ActiveDocument.Unit = cdrMillimeter
    Application.Optimization = True
    
    If tbtnReversePage Then
        signOnPagesNoRevers
    Else
        signOnManyPages
    End If
    
    
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
    ActiveDocument.ClearSelection
End Sub

Private Sub btnPickPlace_Click()
    tbPlateOffset.Enabled = False
    lblPlateOffset.Enabled = False
    btnPickPlace.BackColor = &H8000000D
    btnPickPlace.ForeColor = &H8000000E
    retval = ActiveDocument.GetUserClick(signX, signY, modShift, 100, False, cdrCursorPick)
    customPlace = True
End Sub

Private Sub tbtnHoriz_Click()
    If tbtnHoriz.Value Then
        tbtnVertic.Value = False
    Else
        tbtnVertic.Value = True
    End If
End Sub

Private Sub tbtnVertic_Click()
    If tbtnVertic.Value Then
        tbtnHoriz.Value = False
    Else
        tbtnHoriz.Value = True
    End If
End Sub

Private Sub UserForm_Initialize()
    ActiveDocument.Unit = cdrMillimeter
    tbPageWidth.Value = 497
    tbPageHeight.Value = 347
    tbStartPage.Value = ActivePage.Index
    tbLastPage.Value = ActiveDocument.Pages.Count
    tbStartNumber.Value = 1
    tbPlateOffset.Value = 18
    tbtnReversePage.Value = True
    tbSign.Text = "#0000, 4+4, 347*497, BOSSART 115, спуск $"
    sChar = "$"
    customPlace = False
    offsetBottom = 20
End Sub

Sub signOnManyPages()
    Dim pWidth As Integer, pHeight As Integer
    Dim placeText As String, beginS As String, lastS As String, iChar As Integer
    Dim iPage As Integer, iSpusk As Integer, iEven As Integer
    Dim sign As Shape
    Dim aPage As Page
    
    iSpusk = tbStartNumber.Value
    iEven = 1
    
    iChar = InStr(tbSign.Text, sChar)
    beginS = Left(tbSign.Text, iChar - 1)
    lastS = Mid(tbSign.Text, iChar + 1)
    
    For iPage = tbStartPage.Value To tbLastPage.Value
        Set aPage = ActiveDocument.Pages(iPage)
        If (iEven Mod 2) Then
            placeText = beginS & iSpusk & " лицо" & lastS
        Else
            placeText = beginS & iSpusk & " оборот" & lastS
            iSpusk = iSpusk + 1
        End If
        iEven = iEven + 1
        
        Set sign = aPage.ActiveLayer.CreateArtisticText(0, 0, placeText, , , "Arial", 9, cdrTrue, cdrFalse, , cdrLeftAlignment)
        
        If Not customPlace Then
            signX = (aPage.BoundingBox.CenterX + (tbPageWidth.Value / 2) - (sign.BoundingBox.Height / 2))
            signY = (aPage.BoundingBox.Top - tbPlateOffset.Value - tbPageHeight.Value + offsetBottom)
        End If
        
        sign.PositionX = signX
        sign.PositionY = signY + sign.BoundingBox.Height
        If tbtnVertic Then
            sign.RotationCenterX = sign.BoundingBox.Left
            sign.Rotate (90)
        End If
    Next iPage
End Sub

Sub signOnPagesNoRevers()
    Dim pWidth As Integer, pHeight As Integer
    Dim placeText As String, beginS As String, lastS As String, iChar As Integer
    Dim iPage As Integer, iSpusk As Integer, iEven As Integer
    Dim sign As Shape
    Dim aPage As Page
    
    iSpusk = tbStartNumber.Value
    iEven = 1
    
    iChar = InStr(tbSign.Text, sChar)
    beginS = Left(tbSign.Text, iChar - 1)
    lastS = Mid(tbSign.Text, iChar + 1)
    
    For iPage = tbStartPage.Value To tbLastPage.Value
        Set aPage = ActiveDocument.Pages(iPage)
        placeText = beginS & iSpusk & lastS
        iSpusk = iSpusk + 1
        
        Set sign = aPage.ActiveLayer.CreateArtisticText(0, 0, placeText, , , "Arial", 9, cdrTrue, cdrFalse, , cdrLeftAlignment)
        
        If Not customPlace Then
            signX = (aPage.BoundingBox.CenterX + (tbPageWidth.Value / 2) - (sign.BoundingBox.Height / 2))
            signY = (aPage.BoundingBox.Top - tbPlateOffset.Value - tbPageHeight.Value + offsetBottom)
        End If
        
        sign.PositionX = signX
        sign.PositionY = signY + sign.BoundingBox.Height
        If tbtnVertic Then
            sign.RotationCenterX = sign.BoundingBox.Left
            sign.Rotate (90)
        End If
    Next iPage
End Sub
