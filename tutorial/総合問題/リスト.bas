Attribute VB_Name = "リスト"
Sub リスト参照()
    ActiveWindow.NewWindow
    Windows.Arrange ArrangeStyle:=xlTiled
    Windows("総合問題6.xlsm:1").Activate
    Sheets("リスト").Select
    Windows("総合問題6.xlsm:2").Activate
    ActiveWindow.Zoom = 75
    Range("C15").Select
End Sub
Sub 参照リストを閉じる()
    Windows("総合問題6.xlsm:1").Activate
    ActiveWindow.Close
    ActiveWindow.WindowState = xlMaximized
    ActiveWindow.Zoom = 100
    Range("A1").Select
End Sub
