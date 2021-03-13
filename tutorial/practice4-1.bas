Attribute VB_Name = "STEP2"
 Sub kingaku()
    Dim tanka As Integer
    Dim kazu As Integer
    Dim uriage As Integer
    tanka = Range("C9").Value
    kazu = Range("E9").Value
    uriage = tanka * kazu
    MsgBox uriage
 End Sub
