VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3810
   OleObjectBlob   =   "UserForm1_parsing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim inputStr As String
    Dim c As New Collection
    Dim cTemp As New Collection
    Dim cPages As New Collection
    Dim a As Integer, b As Integer, s As String
    Dim i As Integer, e As Integer
    Dim cObj As Variant
    Dim notEmpty As Boolean
    
    inputStr = TextBox1.Value
    notEmpty = False
    
    For i = 1 To Len(inputStr)
        If isDigit(Mid(inputStr, i, 1)) Then
            a = getDigit(Mid(inputStr, i))
            cTemp.Add a
            i = i + Len(CStr(a)) - 1
            notEmpty = True
        End If
        If Mid(inputStr, i, 1) = "-" Then
            cTemp.Add "-"
            notEmpty = True
        End If
        If Mid(inputStr, i, 1) = "," Or i = Len(inputStr) Then
            If notEmpty Then
                c.Add cTemp
                notEmpty = False
                Set cTemp = New Collection
            End If
        End If
    Next i
    
    For Each cObj In c
        If cObj.Count = 1 And TypeName(cObj.Item(1)) = "Integer" Then
            i = cObj.Item(1)
            If i <= ActiveDocument.Pages.Count Then cPages.Add i
        ElseIf cObj.Count = 2 Then
            If TypeName(cObj.Item(1)) = "String" Then
                For i = 1 To cObj.Item(2)
                    If i <= ActiveDocument.Pages.Count Then cPages.Add i
                Next i
            End If
            If TypeName(cObj.Item(1)) = "Integer" Then
                For i = cObj.Item(1) To ActiveDocument.Pages.Count
                    If i <= ActiveDocument.Pages.Count Then cPages.Add i
                Next i
            End If
        ElseIf cObj.Count = 3 Then
            If TypeName(cObj.Item(1)) = "Integer" And TypeName(cObj.Item(3)) = "Integer" Then
                For i = cObj.Item(1) To cObj.Item(3)
                    If i <= ActiveDocument.Pages.Count Then cPages.Add i
                Next i
            End If
        Else
            MsgBox ("ERROR WTF!" & " | " & cObj.Item(1))
            Exit Sub
        End If
    Next cObj
    testPrint cPages
End Sub

Sub testPrint(c As Collection)
    Dim i As Integer
    ListBox1.Clear
    For i = 1 To c.Count
        ListBox1.AddItem (c.Item(i))
    Next i
End Sub

Function isDigit(str As String) As Boolean
    isDigit = False
    Select Case str
        Case "0"
            isDigit = True
        Case "1"
            isDigit = True
        Case "2"
            isDigit = True
        Case "3"
            isDigit = True
        Case "4"
            isDigit = True
        Case "5"
            isDigit = True
        Case "6"
            isDigit = True
        Case "7"
            isDigit = True
        Case "8"
            isDigit = True
        Case "9"
            isDigit = True
    End Select
End Function

Function getDigit(pStr As String) As Integer
    Dim s As String
    Dim dFind As Boolean
    Dim i As Integer
    dFind = False
    For i = 1 To Len(pStr)
        If isDigit(Mid(pStr, i, 1)) Then
            s = s & Mid(pStr, i, 1)
            dFind = True
        End If
        If Not isDigit(Mid(pStr, i, 1)) And dFind Then
            getDigit = CInt(s)
            Exit Function
        End If
        If i = Len(pStr) And dFind Then
            getDigit = CInt(s)
            Exit Function
        End If
    Next i
End Function

Private Sub UserForm_Click()

End Sub
