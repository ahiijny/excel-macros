Attribute VB_Name = "Module2"
Private Function CanIncrement() As Boolean
    ' Returns true only if the active cell is in a column with "Ch" as the header.
    ' If so, then incrementing the cell will create an archive log.
    
    If IsNumeric(ActiveCell.Value) Then
        If ActiveCell.Offset(1 - ActiveCell.Row).Value = "Ch" Then
            CanIncrement = True
            Exit Function
        End If
    End If
    CanIncrement = False
End Function

Private Function finputdict(finput As String) As Collection
    ' Returns a dictionary giving the column numbers or cell references
    ' for all of the relevant parameters for the Finput sheet
    ' (flat chapter log table)
    
    Dim dict As Collection
    Set dict = New Collection
    
    ' Another option is to hard-code column numbers
    
    With dict
        .Add Application.Match("#", Sheets(finput).Range("A1:K1"), 0), "rowcol"
        .Add Application.Match("Title", Sheets(finput).Range("A1:K1"), 0), "titlecol"
        .Add Application.Match("Fandom", Sheets(finput).Range("A1:K1"), 0), "fandomcol"
        .Add Application.Match("Ch", Sheets(finput).Range("A1:K1"), 0), "chcol"
        .Add Application.Match("Date", Sheets(finput).Range("A1:K1"), 0), "datecol"
        .Add Application.Match("Author", Sheets(finput).Range("A1:K1"), 0), "authorcol"
        .Add Sheets(finput).Cells(1, Application.Match("Last Entry:", Sheets(finput).Range("A1:K1"), 0) + 1), "countref"
    End With
    Set finputdict = dict
    
End Function

Private Function listdict(list As String) As Collection
    ' Returns a dictionary giving the column numbers or cell references
    ' for all of the relevant parameters for the summary / list sheet
    ' (e.g. "Fanfiction", "Books", "Books (archived)", etc.)
    
    Dim dict As Collection
    Set dict = New Collection
    
    ' Another option is to hard-code column numbers
    
    With dict
        .Add Application.Match("Title", Sheets(list).Range("A1:K1"), 0), "titlecol"
        .Add Application.Match("Fandom", Sheets(list).Range("A1:K1"), 0), "fandomcol"
        .Add Application.Match("Ch", Sheets(list).Range("A1:K1"), 0), "chcol"
        .Add Application.Match("Author", Sheets(list).Range("A1:K1"), 0), "authorcol"
        .Add Application.Match("Link", Sheets(list).Range("A1:K1"), 0), "linkcol"
    End With
    Set listdict = dict
End Function

Public Sub PlusPlus()
Attribute PlusPlus.VB_ProcData.VB_Invoke_Func = "p\n14"
    ' Increment the selected cell by one if it is numeric.
    '
    ' If the incremented cell is in the "Ch" column, insert
    ' a title-chapter-date-time-author entry in the Finput sheet
    ' to log the completion of a chapter.
    
    If CanIncrement() Then
        ActiveCell.Value = ActiveCell.Value + 1
        
        Dim finput As String
        finput = "Finput"
        Set fref = finputdict(finput)
        Set lref = listdict(ActiveSheet.Name)
        
        Dim nextrow As Integer
        nextrow = fref.Item("countref").Value + 2
        
        If nextrow >= 0 And nextrow < 65536 Then
            ' Set index and increment ref
            Dim target As Range
            Dim read As Range
            Set target = Sheets(finput).Cells(nextrow, 1)
            Set read = ActiveCell.Offset(0, 1 - ActiveCell.Column)
                        
            ' Set row number, title, chapter, author
            
            target.Offset(0, fref.Item("rowcol") - 1).Value = nextrow - 1
            target.Offset(0, fref.Item("titlecol") - 1).Value = read.Offset(0, lref.Item("titlecol") - 1).Value
            target.Offset(0, fref.Item("chcol") - 1).Value = read.Offset(0, lref.Item("chcol") - 1).Value
            target.Offset(0, fref.Item("authorcol") - 1).Value = read.Offset(0, lref.Item("authorcol") - 1).Value
            
            ' Set date
            
            target.Offset(0, fref.Item("datecol") - 1).Formula = "=TEXT(NOW(),""yyyy-mm-dd  HH:mm"")"
            target.Offset(0, fref.Item("datecol") - 1).Value = target.Offset(0, fref.Item("datecol") - 1).Value
        End If
    End If
End Sub

Public Sub MinusMinus()
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
