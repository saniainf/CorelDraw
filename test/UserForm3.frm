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
Dim cyanColor As New Color, magentaColor As New Color, yellowColor As New Color, blackColor As New Color
Dim whiteColor As New Color
Dim cyan80 As New Color, cyan40 As New Color
Dim magenta80 As New Color, magenta40 As New Color
Dim yellow80 As New Color, yellow40 As New Color
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
    For Each c1 In c
        ListBox2.AddItem c1.Name
    Next c1
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
    cyanColor.CMYKAssign 100, 0, 0, 0
    magentaColor.CMYKAssign 0, 100, 0, 0
    yellowColor.CMYKAssign 0, 0, 100, 0
    blackColor.CMYKAssign 0, 0, 0, 100
    whiteColor.CMYKAssign 0, 0, 0, 0
    cyan80.CMYKAssign 80, 0, 0, 0
    cyan40.CMYKAssign 40, 0, 0, 0
    magenta80.CMYKAssign 0, 80, 0, 0
    magenta40.CMYKAssign 0, 40, 0, 0
    yellow80.CMYKAssign 0, 0, 80, 0
    yellow40.CMYKAssign 0, 0, 40, 0
    black80.CMYKAssign 0, 0, 0, 80
    black40.CMYKAssign 0, 0, 0, 40
    grayBalance.CMYKAssign 38, 26, 26, 0
    
    c.Add cyanColor
    c.Add magentaColor
    c.Add yellowColor
    c.Add blackColor
    c.Add black80
    c.Add black40
    lb1Refresh
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



