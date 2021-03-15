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
Sub datainput()
    Cells(3, 4).Value = InputBox("D3のデータを入力してください", " A店データ入力")
    Cells(3, 6).Value = InputBox("F3のデータを入力してください", " B店データ入力")
    Cells(3, 8).Value = InputBox("H3のデータを入力してください", " C店データ入力")
    Cells(3, 10).Value = InputBox("J3のデータを入力してください", " D店データ入力")
    Cells(3, 12).Value = InputBox("L3のデータを入力してください", " E店データ入力")
End Sub

