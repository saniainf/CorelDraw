VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptions 
   Caption         =   "НАстройки"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2640
   OleObjectBlob   =   "frmOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim allOk As Boolean
    allOk = True
    If CInt(TextBox1.Value) < 6 Then
        MsgBox "Мин высота 6"
        allOk = False
    ElseIf CInt(TextBox2.Value) < 6 Then
        MsgBox "Мин ширина 6"
        allOk = False
    ElseIf CInt(TextBox3.Value) < 5 Then
        MsgBox "Мин количество мин 5"
        allOk = False
    End If
    If allOk Then
       UserForm2.SetOptions TextBox1.Value, TextBox2.Value, TextBox3.Value
    End If
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
