Attribute VB_Name = "Module1"
' 心斎橋 6月'

Sub shinsaibasshi_6()    '累計から日計を導出'
    Dim i As Integer
     
    Dim j As Integer
    For j = 3 To 11 Step 2      ' 列3~11を2列飛ばしで選択
        For i = 3 To 33     '行3~33を選択
             ActiveSheet.Cells(i, j).Value = Abs(ActiveSheet.Cells(i, j + 1).Value - ActiveSheet.Cells(i - 1, j + 1).Value)
             Select Case i
                Case 3
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 4
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 5
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 6
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 7
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 8
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 9
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 10
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 11
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 12
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 13
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 14
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 15
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 16
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 17
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 18
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 19
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 20
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 21
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 22
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 23
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 24
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 25
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 26
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 27
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 28
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 29
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 30
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 31
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 32
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
                Case 33
                    For k = 3 To 11
                        If k Mod 2 = 1 Then
                            Worksheets("6月").Cells(4, k).Value = ActiveSheet.Cells(i, k).Value
                        Else
                            Worksheets("6月").Cells(5, k - 1).Value = ActiveSheet.Cells(i, k).Value
                        End If
                    Next
        Next
    Next

End Sub
Sub clear()
    Range(ActiveSheet.Cells(3, 3), ActiveSheet.Cells(33, 12)).Value = ""
End Sub
