Attribute VB_Name = "STEP7"
Sub Loop1()
    Dim i As Integer
    i = 1
    Range("C7").Select
    Do While ActiveCell.Value <> ""
        MsgBox i & "‰ñ–Ú‚ÌÀs‚Å‚·"
        i = i + 1
        ActiveCell.Offset(0, 1).Select
    Loop
End Sub
Sub Loop2()
    Dim i As Integer
    i = 1
    Range("C14").Select
    Do
        MsgBox i & "‰ñ–Ú‚ÌÀs‚Å‚·"
        i = i + 1
        ActiveCell.Offset(0, 1).Select
    Loop While ActiveCell.Value <> ""
End Sub
