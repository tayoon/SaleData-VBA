Attribute VB_Name = "STEP6"
Sub kaisuu()
    Dim i As Integer
    For i = 1 To 3
        MsgBox i & "âÒñ⁄ÇÃé¿çsÇ≈Ç∑"
    Next
End Sub
Sub zouka()
    Dim i As Integer
    For i = 100 To 300 Step 50
        ActiveCell.Value = i
        ActiveCell.Offset(0, 1).Select
    Next
End Sub
Sub sakujyo()
    Dim i As Integer
    For i = Worksheets.Count To 7 Step -1
        Worksheets(i).Delete
    Next
End Sub
