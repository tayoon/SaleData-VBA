Attribute VB_Name = "Module9"
'  神戸 8月'

Sub kobe_8()    '累計から日計を導出'
    Dim i As Integer
     
    Dim j As Integer
    For j = 3 To 11 Step 2      ' 列3~11を2列飛ばしで選択
        For i = 3 To 33     '行3~33を選択
             ActiveSheet.Cells(i, j).Value = Abs(ActiveSheet.Cells(i, j + 1).Value - ActiveSheet.Cells(i - 1, j + 1).Value)
        Next
    Next
    
    Dim sheetName As String
    Dim month As Integer
    Dim day As String
    sheetName = ActiveSheet.Name
    If sheetName = "心斎橋" Then
        
    ElseIf sheetName = "京都" Then
    
    ElseIf sheetName = "神戸" Then
    
    ElseIf sheetName = "梅田" Then
    End If
    month = 6
    For i = 3 To 33
        day = Cells(i, 1).Value
    
    Worksheets(month + "月").Active
End Sub
Sub clear()
    Range(ActiveSheet.Cells(3, 3), ActiveSheet.Cells(33, 12)).Value = ""
End Sub

