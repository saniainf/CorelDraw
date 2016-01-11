VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Dim WithEvents btn As ToggleButton
Attribute btn.VB_VarHelpID = -1
Dim frm As UserForm2
Dim X As Integer
Dim Y As Integer
 
Public Sub SetHandler(cntrl As UserForm)
    Set frm = cntrl
End Sub

Public Sub SetButton(cntrl As ToggleButton)
    Set btn = cntrl
End Sub

Public Sub SetX(i As Integer)
    X = i
End Sub

Public Sub SetY(i As Integer)
    Y = i
End Sub

Private Sub btn_Click()
    frm.Event_Click Me
End Sub

Private Sub btn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    frm.Event_MouseUp Button, Shift, X, Y, btn
End Sub

Public Function getButton() As ToggleButton
    Set getButton = btn
End Function

Public Function GetX() As Integer
    GetX = X
End Function

Public Function GetY() As Integer
    GetY = Y
End Function
