Attribute VB_Name = "月別販売データ"
Sub 入力()
Attribute 入力.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 入力 Macro
'

'
    ActiveWindow.NewWindow
    Windows.Arrange ArrangeStyle:=xlVertical
    Windows("第5章.xlsm:1").Activate
    Sheets("得意先リスト").Select
    Windows("第5章.xlsm:2").Activate
    Range("C4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "9/1/2016"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "1001"
    Windows("第5章.xlsm:1").Activate
    Sheets("商品リスト").Select
    Windows("第5章.xlsm:2").Activate
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "R101"
    ActiveCell.Offset(0, 4).Range("A1").Select
    ActiveCell.FormulaR1C1 = "10"
    ActiveCell.Offset(1, -7).Range("A1").Select
    Windows("第5章.xlsm:1").Activate
    ActiveWindow.Close
    Range("A1").Select
End Sub
Sub 新規シート()
Attribute 新規シート.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 新規シート Macro
'

'
    Sheets("9月度").Select
    Sheets("9月度").Copy After:=Sheets(4)
    Range("C5:D14,F5:F14,J5:J14").Select
    Range("J5").Activate
    Selection.ClearContents
    Range("A1").Select
End Sub
