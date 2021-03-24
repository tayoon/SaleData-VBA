Attribute VB_Name = "Module3"
Sub 日計8月版()
    Dim i As Integer
    Dim j As Integer
    For j = 3 To 11 Step 2      ' 列3~11を2列飛ばしで選択
        For i = 73 To 103     '行37~103を選択
             Cells(i, j).Value = Abs(Cells(i, j + 1).Value - Cells(i - 1, j + 1).Value)
        Next
    Next
End Sub
Sub clear()
    Range(Cells(73, 3), Cells(103, 12)).Value = ""
End Sub

