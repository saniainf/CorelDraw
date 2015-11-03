Attribute VB_Name = "RecordedMacros"
Sub Macro1()

Dim ss As String, sc As String
Dim startS As String, lastS As String
Dim i As Integer

ss = "#2345, 4+4, 357*589, Атстар$та мат 456, спуск $"
sc = "$"

i = InStr(ss, sc)
startS = Left(ss, i - 1)
lastS = Mid(ss, i + 1)

If (i Mod 2) Then
    MsgBox "False", vbOKOnly, i Mod 2
Else
    MsgBox "True", vbOKOnly, i Mod 2
End If

End Sub

Sub Test()
    Application.ActiveDocument.Unit = cdrMillimeter
    
    Dim ss As String, sc As String
    Dim startS As String, lastS As String
    Dim i As Integer, a As Integer
    Dim s As Shape
    
    Dim retval As Long
    Dim x As Double, y As Double, shift As Long
    
    ss = "#2345, 4+4, 357*589, Атстарта мат 456, спуск $ "
    sc = "$"
    
    i = InStr(ss, sc)
    startS = Left(ss, i - 1)
    lastS = Mid(ss, i + 1)
    a = 1
    
    retval = ActiveDocument.GetUserClick(x, y, shift, 10, True, cdrCursorPick)
    
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
    
End Sub

Sub testfrm()
    UserForm1.Show vbModeles
End Sub

Sub frmUpadate()
    UserForm1.txtbFirst.Text = ActiveSelectionRange.BoundingBox.Height
    UserForm1.txtbSecond.Text = ActiveSelectionRange.BoundingBox.Width
    UserForm1.txtbThird.Text = ActiveSelectionRange.Count
End Sub
