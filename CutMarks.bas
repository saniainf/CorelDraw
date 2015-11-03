Attribute VB_Name = "CutMarks"
Sub CutMarks()
If (Documents.Count = 0) Then Exit Sub
If ActiveSelectionRange.Count <= 0 Then Exit Sub
    frmCutMarks.Show vbModeless
End Sub
