VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMakeOnAllPages 
   Caption         =   "Выполнить на всех страницах"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4545
   OleObjectBlob   =   "frmMakeOnAllPages.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMakeOnAllPages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGroup_Click()
ActiveDocument.BeginCommandGroup "Make on All Pages"
Application.Optimization = True

    Dim aPage As Page
    For Each aPage In ActiveDocument.Pages
        aPage.Activate
        InfUtilits.GroupAll.GroupAll
    Next aPage
    ActiveDocument.Pages(1).Activate
    
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh
ActiveDocument.EndCommandGroup
    Unload Me
End Sub

Private Sub btnOffset10_Click()
ActiveDocument.BeginCommandGroup "Offset Shape 10 mm"
Application.Optimization = True

    Dim aPage As Page
    For Each aPage In ActiveDocument.Pages
        aPage.Activate
        InfUtilits.OffsetAllShapes.OffsetAllShapes10
    Next aPage
    ActiveDocument.Pages(1).Activate
    
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh
ActiveDocument.EndCommandGroup
    Unload Me
End Sub

Private Sub btnOffset12_Click()
ActiveDocument.BeginCommandGroup "Offset Shape 12 mm"
Application.Optimization = True

    Dim aPage As Page
    For Each aPage In ActiveDocument.Pages
        aPage.Activate
        InfUtilits.OffsetAllShapes.OffsetAllShapes12
    Next aPage
    ActiveDocument.Pages(1).Activate
    
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh
ActiveDocument.EndCommandGroup
    Unload Me
End Sub

Private Sub btnPowerClip_Click()
ActiveDocument.BeginCommandGroup "Place to Power Clip"
Application.Optimization = True

    Dim aPage As Page
    For Each aPage In ActiveDocument.Pages
        aPage.Activate
        InfUtilits.PlaceAllToPowerClip.PlaceAllToPowerClip
    Next aPage
    ActiveDocument.Pages(1).Activate
    
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh
ActiveDocument.EndCommandGroup
    Unload Me
End Sub

Private Sub btnPrintMarksR5_Click()
ActiveDocument.BeginCommandGroup "Create R5 Print Marks"
Application.Optimization = True
    
    Dim aPage As Page
    For Each aPage In ActiveDocument.Pages
        aPage.Activate
        InfUtilits.PrintMarksR5.PrintMarksR5
    Next aPage
    ActiveDocument.Pages(1).Activate
    
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh
ActiveDocument.EndCommandGroup
    Unload Me
End Sub

Private Sub btnPrintMarksR7_Click()
ActiveDocument.BeginCommandGroup "Create R5 Print Marks"
Application.Optimization = True

    Dim aPage As Page
    For Each aPage In ActiveDocument.Pages
        aPage.Activate
        InfUtilits.PrintMarksR7.PrintMarksR7
    Next aPage
    ActiveDocument.Pages(1).Activate
    
Application.Optimization = False
ActiveWindow.Refresh
Application.Refresh
ActiveDocument.EndCommandGroup
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
