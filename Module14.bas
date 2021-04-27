Attribute VB_Name = "Module14"
 'Clear'
 
 Sub monthClear()
    For Each c In ActiveSheet.Range("C1:G403").Cells
        If c.Interior.Color = RGB(0, 0, 0) Then
           c.Value = ""
        End If
    Next
 End Sub
