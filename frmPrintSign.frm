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

Private Sub btnMake_Click()
    Dim pWidth As Integer, pHeight As Integer
    Dim placeText As String
    Dim iPage As Integer, iSpusk As Integer, iEven As Integer
    
    
End Sub

Private Sub btnPickPlace_Click()
    tbPlateOffset.Enabled = False
    lblPlateOffset.Enabled = False
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
    tbPageWidth.Value = 497
    tbPageHeight.Value = 347
    tbStartPage.Value = ActivePage.Index
    tbLastPage.Value = ActiveDocument.Pages.Count
    tbStartNumber.Value = 1
    tbPlateOffset.Value = 18
    tbtnReversePage.Value = True
    tbSign.Text = "#0000, 4+4, 347*497, BOSSART 115, спуск $"
End Sub
