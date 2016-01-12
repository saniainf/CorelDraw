VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "ÌÈíåð"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13485
   OleObjectBlob   =   "UserForm2mine.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Col As New Collection
Dim matrix() As Integer
Dim w As Integer, h As Integer
Dim mineCount As Integer
Dim cellSize As Integer

Private Sub CommandButton2_Click()
    frmOptions.Show
End Sub

Public Sub SetOptions(e As String, i As String, a As String)
    w = CInt(e) - 1
    h = CInt(i) - 1
    mineCount = CInt(a) - 1
    ReDim matrix(w, h)
    newGame
End Sub

Private Sub UserForm_Initialize()
    cellSize = 24
    w = 5   '-1
    h = 5   '-1
    mineCount = 4   '-1
    ReDim matrix(w, h)
    newGame
End Sub

Private Sub CommandButton1_Click()
    newGame
End Sub

Public Function newGame()
    Dim i As Integer, e As Integer
    Dim X As Integer, Y As Integer
    Dim n As Integer
    Dim tb As ToggleButton
    Dim wrp As clsEventWrapper
    If Col.Count > 1 Then
        For i = 0 To Col.Count - 1
            Me.Controls.Remove "tb" & i
        Next i
    End If
    Set Col = New Collection
    n = 0
    For i = 0 To h
        For e = 0 To w
            Set wrp = New clsEventWrapper
            wrp.SetHandler Me
            Set tb = Me.Controls.Add("Forms.ToggleButton.1", "tb" & n)
            tb.height = cellSize
            tb.width = cellSize
            tb.Left = cellSize * e
            tb.Top = cellSize * i
            tb.TabStop = False
            tb.Value = False
            wrp.SetButton tb
            wrp.SetX e
            wrp.SetY i
            Col.Add wrp
            n = n + 1
        Next e
    Next i
    Randomize
    clearMatrix
    For i = 0 To mineCount
        X = CInt(Int((w + 1) * Rnd()))
        Y = CInt(Int((h + 1) * Rnd()))
        If matrix(X, Y) = 66 Then
            i = i - 1
        Else
            matrix(X, Y) = 66
        End If
    Next i
    countMinesAround
    formOptions
End Function

Sub formOptions()
    Me.width = cellSize * (w + 1) + 4
    Me.height = cellSize * (h + 1) + 52
    CommandButton1.Left = 3
    CommandButton1.Top = Me.height - CommandButton1.height * 2
    CommandButton2.Left = Me.width - CommandButton2.width - 8
    CommandButton2.Top = Me.height - CommandButton1.height * 2
End Sub
Public Function Event_Click(wrp As clsEventWrapper)
    If wrp.getButton.Caption = "" Then
        Dim e As Integer, i As Integer
        e = wrp.GetX
        i = wrp.GetY
        wrp.getButton.Locked = True
        If matrix(e, i) = 0 Then
            openEmpty e, i
        ElseIf matrix(e, i) = 66 Then
            wrp.getButton.Caption = "X"
            wrp.getButton.Font.Bold = True
            wrp.getButton.Font.Size = 12
            wrp.getButton.ForeColor = &HFF&
            gameOver
        Else
            wrp.getButton.Caption = matrix(e, i)
            wrp.getButton.Font.Bold = True
            wrp.getButton.Font.Size = 12
            Select Case matrix(e, i)
                Case 1
                    wrp.getButton.ForeColor = &HFF0000
                Case 2
                    wrp.getButton.ForeColor = &HC000&
                Case 3
                    wrp.getButton.ForeColor = &HC0&
                Case 4
                    wrp.getButton.ForeColor = &H800000
                Case 5
                    wrp.getButton.ForeColor = &H4080&
                Case 6
                    wrp.getButton.ForeColor = &H808000
                Case 7
                    wrp.getButton.ForeColor = &H0&
                Case 8
                    wrp.getButton.ForeColor = &H808080
            End Select
        End If
    checkWin
    Else
        wrp.getButton.Value = False
    End If
End Function
Sub gameOver()
    MsgBox "GAME OVER"
    newGame
End Sub

Sub checkWin()
    Dim i As Integer, e As Integer
    Dim tb As ToggleButton
    Dim win As Boolean
    win = True
    For i = 0 To h
        For e = 0 To w
            Set tb = Col((i * (w + 1) + e) + 1).getButton
            If tb.Locked = False And Not matrix(e, i) = 66 Then
                win = False
            End If
        Next e
    Next i
    If win Then
        MsgBox "WIN"
        newGame
    End If
End Sub

Sub openEmpty(e As Integer, i As Integer)
    Dim e2 As Integer, i2 As Integer
    Dim tb As ToggleButton
    For i2 = i - 1 To i + 1
        For e2 = e - 1 To e + 1
            If i2 >= 0 And i2 <= h And e2 >= 0 And e2 <= w Then
                Set tb = Col((i2 * (w + 1) + e2) + 1).getButton
                If tb.Caption = "" Then
                    If tb.Locked = False Then
                        tb.Locked = True
                        tb.Value = True
                    End If
                End If
            End If
        Next e2
    Next i2
End Sub

Public Function Event_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, cntrl As ToggleButton)
    If Button = 2 Then
        If cntrl.Caption = "" Then
            cntrl.Caption = "P"
            cntrl.Font.Bold = True
            cntrl.Font.Size = 12
            cntrl.ForeColor = &HFF&
        Else
            cntrl.Caption = ""
        End If
    End If
End Function

Sub clearMatrix()
    Dim i As Integer, e As Integer
    For i = 0 To h
        For e = 0 To w
            matrix(e, i) = 0
        Next e
    Next i
End Sub

Sub countMinesAround()
    Dim i As Integer, e As Integer
    Dim i2 As Integer, e2 As Integer
    Dim mCount As Integer
    For i = 0 To h
        For e = 0 To w
            mCount = 0
            For i2 = i - 1 To i + 1
                For e2 = e - 1 To e + 1
                    If i2 >= 0 And i2 <= h And e2 >= 0 And e2 <= w Then
                        If matrix(e2, i2) = 66 Then mCount = mCount + 1
                    End If
                Next e2
            Next i2
            If mCount > 0 And Not matrix(e, i) = 66 Then matrix(e, i) = mCount
        Next e
    Next i
End Sub

