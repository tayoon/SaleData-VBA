Attribute VB_Name = "STEP4"
 Sub shiken1()
    Dim tensuu As Integer
    Dim goukaku As Integer
    tensuu = Range("C8").Value
    goukaku = Range("E8").Value
    If tensuu >= goukaku Then
        MsgBox "‡Ši‚Å‚·"
    End If
 End Sub
 
