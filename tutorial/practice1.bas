Attribute VB_Name = "Module1"
Sub ﾌｫﾝﾄの変更()
    With Selection.Font
        .Size = 14
        .ThemeColor = xlThemeColorLight1
    End With
End Sub
Sub フォントの変更2()
    With Selection.Font
        .Size = 20
        .ThemeColor = 6
    End With
End Sub
Sub 集計()
Attribute 集計.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 集計 Macro
'

'
    Range("C8").Select
    ActiveWorkbook.Worksheets("記録の練習2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("記録の練習2").Sort.SortFields.Add2 Key:=Range("C8:C30") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("記録の練習2").Sort
        .SetRange Range("B9:H30")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B8").Select
    Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(7), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    Range("A1").Select
End Sub
