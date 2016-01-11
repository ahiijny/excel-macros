Attribute VB_Name = "Module2"
Sub PlusPlus()
Attribute PlusPlus.VB_ProcData.VB_Invoke_Func = "p\n14"
    ' Increment the selected cell by one if it is numeric.
    '
    ' If the incremented cell is in the "Ch" column, insert
    ' a title-chapter-date-time-author entry in the Finput sheet
    ' to log the completion of a chapter.
    
    If IsNumeric(ActiveCell.Value) Then
        ActiveCell.Value = ActiveCell.Value + 1
        If ActiveCell.Offset(1 - ActiveCell.Row).Value = "Ch" Then
            Dim ref As Range
            Set ref = Sheets("Finput").Range("H1")
            If IsNumeric(ref.Value) And ref.Value >= 0 And ref.Value < 65536 Then
                ' Set index and increment ref
                Dim target As Range
                Set target = Sheets("Finput").Range("A" & (ref.Value + 2))
                ref.Value = ref.Value + 1
                target.Value = ref.Value
                ' Set title, chapter, author
                Dim read As Range
                Set read = ActiveCell.Offset(0, 1 - ActiveCell.Column)
                Dim i As Integer
                For i = 0 To 9
                    Dim str As String
                    Let str = read.Offset(1 - ActiveCell.Row, i).Value
                    If str = "Title" Then
                        target.Offset(0, 1).Value = read.Offset(0, i).Value
                    End If
                    If str = "Ch" Then
                        target.Offset(0, 2).Value = read.Offset(0, i).Value
                    End If
                    If str = "Author" Then
                        target.Offset(0, 4).Value = read.Offset(0, i).Value
                    End If
                Next i
                ' Set date
                
                target.Offset(0, 3).FormulaR1C1 = "=TEXT(TODAY(),""yyyy-mm-dd"") & "" "" & TEXT(NOW(),""HH:mm"")"
                target.Offset(0, 3).Value = target.Offset(0, 3).Value
                
            End If
        End If
    End If
End Sub
Sub MinusMinus()
Attribute MinusMinus.VB_ProcData.VB_Invoke_Func = "n\n14"
' Decrease the value of the selected cell by one if it is numeric.
' If the result is negative, resets the cell value to zero.

    If IsNumeric(ActiveCell.Value) Then
        ActiveCell.Value = ActiveCell.Value - 1
        If ActiveCell.Value < 0 Then
            ActiveCell.Value = 0
        End If
    End If
End Sub
