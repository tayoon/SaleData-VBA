Attribute VB_Name = "Module1"
Sub 日計6月版()
    Dim i As Integer
    Dim j As Integer
    For j = 3 To 11 Step 2      ' 列3~11を2列飛ばしで選択
        For i = 3 To 33     '行3~33を選択
             ActiveSheet.Cells(i, j).Value = Abs(ActiveSheet.Cells(i, j + 1).Value - ActiveSheet.Cells(i - 1, j + 1).Value)
        Next
    Next
End Sub
Sub clear()
    Range(ActiveSheet.Cells(3, 3), ActiveSheet.Cells(33, 12)).Value = ""
End Sub
