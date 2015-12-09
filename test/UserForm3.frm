VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6990
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New Collection
Dim cColor As New Collection
Dim cyanColor As New Color, magentaColor As New Color, yellowColor As New Color, blackColor As New Color
Dim whiteColor As New Color
Dim cyan80 As New Color, cyan40 As New Color
Dim black80 As New Color, black40 As New Color
Dim grayBalance As New Color
Dim pS As Integer

Private Sub CommandButton1_Click()
    Dim cColor As New Color
    If cColor.UserAssignEx Then
        If ListBox1.ListIndex = -1 Then
            ListBox1.ListIndex = ListBox1.ListCount - 1
        End If
        c.Add cColor, , ListBox1.ListIndex + 1
        lb1Refresh
    End If
End Sub

Private Sub CommandButton2_Click()
    ListBox2.Clear
    'ActiveShape.Fill.ApplyUniformFill CreateSpotColor(ActiveSelection.Fill.UniformColor.PaletteIdentifier, ActiveSelection.Fill.UniformColor.SpotColorID, 50)
    'Dim x As Integer
    'For i = 0 To 20
        'x = (i Mod 3) + 1
        'ListBox2.AddItem x
    'Next i
    For Each cc3 In cColor
        'ListBox2.AddItem TypeName(cc3)
        If TypeName(cc3) = "IDrawColor" Then
            ListBox2.AddItem cc3.Name
        End If
        If TypeName(cc3) = "String" Then
            ListBox2.AddItem cc3
        End If
    Next cc3
End Sub

Private Sub CommandButton3_Click()
    If ListBox1.ListCount = 1 Then Exit Sub
    If ListBox1.ListIndex = -1 Then
        ListBox1.ListIndex = ListBox1.ListCount - 1
    End If
    c.Remove ListBox1.ListIndex + 1
    lb1Refresh
End Sub

Private Sub CommandButton4_Click()
    If ListBox1.ListIndex = -1 Then
        ListBox1.ListIndex = ListBox1.ListCount - 1
    End If
    c.Add black80, , ListBox1.ListIndex + 1
    lb1Refresh
End Sub

Private Sub CommandButton5_Click()
    Dim x As Double, spaceWidth As Double, barWidth As Double
    Dim i As Integer
    Dim sBar As Shape
    spaceWidth = 0.3
    barWidth = 3.9
    x = 0
    For i = 1 To 16
        'color bar
        For a = 1 To 8 \ c.Count
            For Each c3 In c
                Set sBar = ActiveLayer.CreateRectangle(x, barWidth, x + barWidth, 0)
                sBar.Outline.SetNoOutline
                sBar.Fill.UniformColor = c3
                x = x + barWidth
            Next c3
        Next a
        'white bar
        For a = 0 To 8 Mod c.Count - 1
            Set sBar = ActiveLayer.CreateRectangle(x, barWidth, x + barWidth, 0)
            sBar.Outline.SetNoOutline
            sBar.Fill.ApplyNoFill
            x = x + barWidth
        Next a
        'white space
        Set sBar = ActiveLayer.CreateRectangle(x, barWidth, x + spaceWidth, 0)
        sBar.Outline.SetNoOutline
        sBar.Fill.ApplyNoFill
        x = x + spaceWidth
    Next i
End Sub

Private Sub ListBox1_Change()
    If ListBox1.ListCount <= 0 Or ListBox1.ListIndex < 0 Then Exit Sub
    pS = ListBox1.ListIndex
    Label1.Caption = "scroll min " & ScrollBar1.Min & vbCrLf _
                & "scroll max " & ScrollBar1.Max & vbCrLf _
                & "list index " & ListBox1.ListIndex & vbCrLf _
                & "list count " & ListBox1.ListCount
End Sub

Private Sub ListBox1_Click()
    If ListBox1.ListCount <= 0 Or ListBox1.ListIndex < 0 Then Exit Sub
    ScrollBar1.Value = ListBox1.ListIndex
End Sub

Private Sub ScrollBar1_Change()
    If ListBox1.ListCount <= 0 Or ListBox1.ListIndex < 0 Then Exit Sub
    
    Dim mColor As New Color
    Set mColor = c.Item(pS + 1)
    
    If pS < ScrollBar1.Value Then
        Label4.Caption = "Down"
        c.Remove pS + 1
        c.Add mColor, , , pS + 1
        ListBox1.ListIndex = ScrollBar1.Value
        lb1Refresh
    End If
    
    If pS > ScrollBar1.Value Then
        Label4.Caption = "Up"
        c.Remove pS + 1
        c.Add mColor, , pS
        ListBox1.ListIndex = ScrollBar1.Value
        lb1Refresh
    End If
    
    Label2.Caption = "prev val " & pS & vbCrLf & "color " & mColor.Name
    Label3.Caption = "scroll val " & ScrollBar1.Value
    
End Sub

Private Sub UserForm_Initialize()
    Application.ActiveDocument.Unit = cdrMillimeter
    cyanColor.CMYKAssign 100, 0, 0, 0
    magentaColor.CMYKAssign 0, 100, 0, 0
    yellowColor.CMYKAssign 0, 0, 100, 0
    blackColor.CMYKAssign 0, 0, 0, 100
    whiteColor.CMYKAssign 0, 0, 0, 0
    cyan80.CMYKAssign 80, 0, 0, 0
    cyan40.CMYKAssign 40, 0, 0, 0
    black80.CMYKAssign 0, 0, 0, 80
    black40.CMYKAssign 0, 0, 0, 40
    grayBalance.CMYKAssign 38, 26, 26, 0
    
    c.Add cyanColor
    c.Add magentaColor
    c.Add yellowColor
    c.Add blackColor
    lb1Refresh
    
    Dim str1 As String, str2 As String
    
    str1 = "string1"
    str2 = "string2"
    
    cColor.Add cyanColor
    cColor.Add magentaColor
    cColor.Add yellowColor
    cColor.Add str1
    cColor.Add str2
End Sub

Sub lb1Refresh()
    Dim s As Integer
    s = ListBox1.ListIndex
    ListBox1.Clear
    For Each c2 In c
        ListBox1.AddItem c2.Name
    Next c2
    
    If s >= ListBox1.ListCount Then
        s = ListBox1.ListCount - 1
    End If
    ListBox1.ListIndex = s
    
    If ListBox1.ListIndex = -1 Then
        ListBox1.ListIndex = ListBox1.ListCount - 1
    End If
    
    ScrollBar1.Min = 0
    ScrollBar1.Max = ListBox1.ListCount - 1
End Sub



