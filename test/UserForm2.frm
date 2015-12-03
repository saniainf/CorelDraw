VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim spotColor As New Collection

Private Sub cmdAddColor_Click()
    Dim cColor As New Color
    If cColor.UserAssignEx Then
        spotColor.Add cColor
        lbSpotColors.AddItem cColor.Name
    End If
End Sub

Private Sub cmdCancel_Click()

End Sub

Private Sub cmdMake_Click()
    ActiveDocument.Unit = cdrMillimeter
    Application.Optimization = True
    
    Dim itemColorBar As Shape
    Dim colorBar As ShapeRange, finalColorBar As ShapeRange, srD As ShapeRange
    Dim printMarksPath As String
    
    printMarksPath = (UserDataPath & "printMarks\")
    
    ActiveLayer.Import (printMarksPath & "colorBarR5.cdr")
    Set colorBar = ActiveSelectionRange
    
    If spotColor.Count < 5 Then
        Set colorBar = CMYK_4Spot(colorBar)
    End If
    
    ActiveDocument.ClearSelection
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Public Function CMYK_4Spot(sr As ShapeRange) As Shape
    Dim cyanColor As New Color, magentaColor As New Color, yellowColor As New Color, blackColor As New Color
    Dim whiteColor As New Color
    Dim cyan80 As New Color, cyan40 As New Color
    Dim magenta80 As New Color, magenta40 As New Color
    Dim yellow80 As New Color, yellow40 As New Color
    Dim black80 As New Color, black40 As New Color
    
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
    
    
    sr.UngroupAll
    
    If Not tbtnCyan.Value Then
        For Each s1 In sr
            If s1.Fill.UniformColor.IsSame(cyanColor) Or s1.Fill.UniformColor.IsSame(cyan80) Or s1.Fill.UniformColor.IsSame(cyan40) Then
                s1.Fill.ApplyUniformFill whiteColor
            End If
        Next s1
    End If
    
    If Not tbtnMagenta.Value Then
        For Each s1 In sr
            If s1.Fill.UniformColor.IsSame(magentaColor) Or s1.Fill.UniformColor.IsSame(magenta80) Or s1.Fill.UniformColor.IsSame(magenta40) Then
                s1.Fill.ApplyUniformFill whiteColor
            End If
        Next s1
    End If
    
    If Not tbtnYellow.Value Then
        For Each s1 In sr
            If s1.Fill.UniformColor.IsSame(yellowColor) Or s1.Fill.UniformColor.IsSame(yellow80) Or s1.Fill.UniformColor.IsSame(yellow40) Then
                s1.Fill.ApplyUniformFill whiteColor
            End If
        Next s1
    End If
    
    If Not tbtnKey.Value Then
        For Each s1 In sr
            If s1.Fill.UniformColor.IsSame(blackColor) Or s1.Fill.UniformColor.IsSame(black80) Or s1.Fill.UniformColor.IsSame(black40) Then
                s1.Fill.ApplyUniformFill whiteColor
            End If
        Next s1
    End If
    
End Function

Private Sub UserForm_Click()

End Sub
