Attribute VB_Name = "Module2"
Sub 日計7月版()
    Dim i As Integer
    Dim j As Integer
    For j = 3 To 11 Step 2      ' 列3~11を2列飛ばしで選択
        For i = 38 To 68     '行38~68を選択
             ActiveSheet.Cells(i, j).Value = Abs(ActiveSheet.Cells(i, j + 1).Value - ActiveSheet.Cells(i - 1, j + 1).Value)
        Next
    Next
End Sub
Sub clear()
    Range(ActiveSheet.Cells(38, 3), ActiveSheet.Cells(68, 12)).Value = ""
End Sub

