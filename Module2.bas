Attribute VB_Name = "Module2"
Sub 日計7月版()
    Dim i As Integer
    Dim j As Integer
    For j = 3 To 11 Step 2      ' 列3~11を2列飛ばしで選択
        For i = 38 To 68     '行38~68を選択
             Cells(i, j).Value = Abs(Cells(i, j + 1).Value - Cells(i - 1, j + 1).Value)
        Next
    Next
End Sub
Sub clear()
    Range(Cells(38, 3), Cells(68, 12)).Value = ""
End Sub

