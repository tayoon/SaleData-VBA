Attribute VB_Name = "Module1"
Sub day()
    Dim i As Integer
    Dim j As Integer
    For j = 3 To 11 Step 2
        For i = 3 To 33
             Cells(i, j).Value = Abs(Cells(i, j + 1).Value - Cells(i - 1, j + 1).Value)
        Next
    Next
End Sub
Sub clear()
    Range(Cells(3, 3), Cells(33, 12)).Value = ""
End Sub
