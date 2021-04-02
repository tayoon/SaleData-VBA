Attribute VB_Name = "Module3"
' 心斎橋 8月'

Sub shinsaibashi_8()
    Dim i As Integer
    Dim j As Integer
    For j = 3 To 11 Step 2      ' 列3~11を2列飛ばしで選択
        For i = 73 To 103     '行37~103を選択
             ActiveSheet.Cells(i, j).Value = Abs(ActiveSheet.Cells(i, j + 1).Value - ActiveSheet.Cells(i - 1, j + 1).Value)
        Next
    Next
End Sub
Sub clear()
    Range(ActiveSheet.Cells(73, 3), ActiveSheet.Cells(103, 12)).Value = ""
End Sub

