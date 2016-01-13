VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLEdit 
   Caption         =   "LEdit"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7335
   OleObjectBlob   =   "frmLEdit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCreate_Click()
    Dim s As Shape
    Dim aSel As ShapeRange
    Dim width As Integer, height As Integer
    Dim cellSize As Integer
    Dim e As Integer, i As Integer
    
    width = CInt(tbWidth.Value)
    height = CInt(tbHeight.Value)
    cellSize = CInt(tbCellSize.Value)
    
    Set aSel = ActivePage.Shapes.All
    tbMap.Value = ""
    
    For Each s In aSel
        e = s.BoundingBox.Left / cellSize
        i = s.BoundingBox.Bottom / cellSize
        tbMap.Value = tbMap.Value & e & "," & i & ","
    Next s
End Sub
