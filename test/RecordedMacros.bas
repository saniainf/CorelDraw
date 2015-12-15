Attribute VB_Name = "RecordedMacros"
Sub Macro1()
    ActiveDocument.BeginCommandGroup "gr1"
    Application.Optimization = True
    Dim s1 As Shape
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    Set s1 = sr.FirstShape.Duplicate
    For i = 2 To sr.Count
        Dim s2 As Shape
        Set s2 = s1.Weld(sr.Item(i), False, False)
        Set s1 = s2
    Next i
    sr.FirstShape.Delete
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
    ActiveDocument.EndCommandGroup
End Sub

Sub Test1()
 Dim d As Document
 Dim p As Page
 Dim c As Color
 Dim s As Shape
 Const sx As Double = 0.5
 Const sy As Double = 0.5
 Dim x As Double, y As Double
 Dim MaxX As Long, nx As Long
 Dim MaxY As Long, ny As Long
 x = 0
 y = 0
 nx = 0
 Set d = CreateDocument
 d.Unit = cdrInch
 Set p = d.ActivePage
 MaxX = CLng(p.SizeWidth / sx)
 MaxY = CLng(p.SizeHeight / sy)
 For Each c In ActivePalette.Colors
  Set s = p.ActiveLayer.CreateRectangle(x, y, x + sx, y + sy)
  s.Fill.ApplyUniformFill c
  x = x + sx
  nx = nx + 1
  If nx = MaxX Then
   nx = 0
   x = 0
   y = y + sy
   ny = ny + 1
   If ny = MaxY Then
    ny = 0
    y = 0
    Set p = d.AddPages(1)
   End If
  End If
 Next c
End Sub

Sub Test2()
 Dim c1 As Color
 Dim c2 As Color
 Dim c3 As Color
 
 Set c1 = ActiveSelection.Fill.UniformColor
 Set c2 = CreateCMYKColor(0, 0, 0, 0)
 'Set c3 = CreateSpotColor.SpotAssign(ActiveSelection.Fill.UniformColor.PaletteIdentifier, ActiveSelection.Fill.UniformColor.SpotColorID, 50)
 'c1.BlendWith c2, 80
 
 ActiveShape.Fill.ApplyUniformFill CreateSpotColor(ActiveSelection.Fill.UniformColor.PaletteIdentifier, ActiveSelection.Fill.UniformColor.SpotColorID, 50)
End Sub

Sub Test3()
 Dim d As Document
 Dim p As Page
 Dim s As Shape
 Dim msg As String
 Dim num As Long
 Dim NoFill As Long, Uniform As Long
 Dim Fountain As Long, Pattern As Long
 Dim Texture As Long, PostScript As Long
 NoFill = 0
 Uniform = 0
 Fountain = 0
 Pattern = 0
 Texture = 0
 PostScript = 0
 num = 0
 Set d = ActiveDocument
 For Each p In d.Pages
  For Each s In p.Shapes
   Select Case s.Fill.Type
    Case cdrNoFill
     NoFill = NoFill + 1
    Case cdrUniformFill
     Uniform = Uniform + 1
    Case cdrFountainFill
     Fountain = Fountain + 1
    Case cdrPatternFill
     Pattern = Pattern + 1
    Case cdrTextureFill
     Texture = Texture + 1
    Case cdrPostscriptFill
     PostScript = PostScript + 1
   End Select
   num = num + 1
  Next s
 Next p
 msg = "The document contains " & num & " shapes with:" & vbCr
 msg = msg & "No fill: " & NoFill & vbCr
 msg = msg & "Uniform fill: " & Uniform & vbCr
 msg = msg & "Fountain fill: " & Fountain & vbCr
 msg = msg & "Pattern fill: " & Pattern & vbCr
 msg = msg & "Texture fill: " & Texture & vbCr
 msg = msg & "PostScript fill: " & PostScript
 MsgBox msg, vbInformation, "Statistics"
End Sub

Sub Test4()
 Dim s As Shape
 For Each s In ActivePage.Shapes
  If s.Fill.Type = cdrPostscriptFill Then
   s.Fill.PostScript.Select "Bricks"
  End If
 Next s
End Sub

Sub Test5()
 Dim s As Shape
 Dim cc As FountainColor
 For Each s In ActivePage.Shapes
  Select Case s.Fill.Type
   Case cdrUniformFill
    s.Fill.UniformColor.ConvertToGray
   Case cdrFountainFill
    s.Fill.Fountain.StartColor.ConvertToGray
    s.Fill.Fountain.EndColor.ConvertToGray
    For Each cc In s.Fill.Fountain.Colors
     cc.Color.ConvertToGray
    Next cc
  End Select
 Next s
End Sub

Sub doWeildAllShape()
   Dim s As Shape, con As Integer
    ActiveDocument.BeginCommandGroup "kkk"
    On Error GoTo ErrHandler
    ActiveDocument.Unit = cdrPixel
    ActiveLayer.Shapes.All.CreateSelection
    Set sh = ActiveDocument.SelectionRange
    con = sh.Count
    If con < 2 Then
        Exit Sub
    End If
    For i = 1 To con - 1
        Set s = ActiveLayer.Shapes(1).Weld(ActiveLayer.Shapes(2), False, False)
        ActiveLayer.Shapes(1).Curve.Nodes.All.AutoReduce 1
    Next i

ExitSub:
        ActiveDocument.EndCommandGroup
    Exit Sub
ErrHandler:
    MsgBox "Error occured: " & Err.Description
    ActiveDocument.Undo
    Resume ExitSub
End Sub


Sub Macro2()
    ' Recorded 14.12.2015
    Dim s1 As Shape
    Set s1 = ActiveLayer.CreateArtisticText(2.691827, 3.069563, "Text")
    s1.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
    s1.Outline.SetNoOutline
    ' Recording of this command is not supported: TextUndoRedo
    s1.ConvertToCurves
    s1.Delete
End Sub
