VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7470
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public retval As Long
Public x As Double, y As Double, Shift As Long

Private Sub btn1_Click()
Application.ActiveDocument.Unit = cdrMillimeter
    
    Dim ss As String, sc As String
    Dim startS As String, lastS As String
    Dim i As Integer, a As Integer
    Dim s As Shape
    

    
    ss = "#2345, 4+4, 357*589, Атстарта мат 456, спуск $ "
    sc = "$"
    
    i = InStr(ss, sc)
    startS = Left(ss, i - 1)
    lastS = Mid(ss, i + 1)
    a = 1
    
    For i = 1 To 20
        If i Mod 2 Then
            Set s = ActiveLayer.CreateArtisticText(x, y + i * 5, startS & a & lastS & " лицо", , , "Arial", 9, cdrTrue, cdrFalse, , cdrLeftAlignment)
            s.Fill.UniformColor.RegistrationAssign
        Else
            Set s = ActiveLayer.CreateArtisticText(x, y + i * 5, startS & a & lastS & " оборот", , , "Arial", 9, cdrTrue, cdrFalse, , cdrLeftAlignment)
            s.Fill.UniformColor.RegistrationAssign
            a = a + 1
        End If
    Next i
    
    Unload Me
End Sub

Private Sub btn2_Click()
    UserForm1.txtbFirst.Text = ActiveSelectionRange.BoundingBox.Height
    UserForm1.txtbSecond.Text = ActiveSelectionRange.BoundingBox.Width
    UserForm1.txtbThird.Text = ActiveSelectionRange.Count
End Sub

Private Sub btn3_Click()
 MsgBox UserDataPath
End Sub

Private Sub cb1_Change()
    If cb1.Value = True Then
        cb2.Enabled = False
        cb3.Enabled = False
    ElseIf cb1.Value = fasle Then
        cb2.Enabled = True
        cb3.Enabled = True
    End If
End Sub

Private Sub tb1_Change()
    If (tb1.Value = True) Then
        tb2.Enabled = False
    ElseIf (tb1.Value = False) Then
        tb2.Enabled = True
    End If
End Sub


Private Sub tb2_Change()
    If (tb2.Value = True) Then
        tb1.Enabled = False
    ElseIf (tb2.Value = False) Then
        tb1.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()

End Sub


