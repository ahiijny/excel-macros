Private Function CanIncrement() As Boolean
    ' Returns true only if the active cell is numeric and in a column with "Ep" as the header.
    
    If IsNumeric(ActiveCell.Value) Then
        If ActiveCell.Offset(1 - ActiveCell.Row).Value = "Ep" Then
            CanIncrement = True
            Exit Function
        End If
    End If
    CanIncrement = False
End Function

Private Function GetLogColumnsInSheet(logsheet As String) As Collection
    ' Returns a dictionary giving the column numbers or cell references
    ' for all of the relevant columns for given log sheet
    ' (expected to be a flat episode log table)
    
    Dim dict As Collection
    Set dict = New Collection
    
    ' Another option is to hard-code column numbers
    
    With dict
        .Add Application.Match("Date", Sheets(logsheet).Range("A1:Z1"), 0), "date"
        .Add Application.Match("Title", Sheets(logsheet).Range("A1:Z1"), 0), "title"
        .Add Application.Match("S", Sheets(logsheet).Range("A1:Z1"), 0), "season"
        .Add Application.Match("Ep", Sheets(logsheet).Range("A1:Z1"), 0), "episode"
        .Add Application.Match("Subtitle", Sheets(logsheet).Range("A1:Z1"), 0), "subtitle"
        .Add Sheets(logsheet).Cells(1, Application.Match("Last Entry:", Sheets(logsheet).Range("A1:Z1"), 0) + 1), "count"
    End With
    Set GetLogColumnsInSheet = dict
    
End Function

Private Function GetSeriesColumnsInSheet(seriessheet As String) As Collection
    ' Returns a dictionary giving the column numbers or cell references
    ' for all of the relevant parameters for the anime / cartoon / TV series sheet
    ' (e.g. "Anime", "Cartoons", "Live-Action TV", etc.)
    
    Dim dict As Collection
    Set dict = New Collection
    
    ' Another option is to hard-code column numbers
    
    With dict
        .Add Application.Match("Studio", Sheets(seriessheet).Range("A1:Z1"), 0), "studio"
        .Add Application.Match("Translation", Sheets(seriessheet).Range("A1:Z1"), 0), "translation"
        .Add Application.Match("Title", Sheets(seriessheet).Range("A1:Z1"), 0), "title"
        .Add Application.Match("S", Sheets(seriessheet).Range("A1:Z1"), 0), "season"
        .Add Application.Match("Ep", Sheets(seriessheet).Range("A1:Z1"), 0), "episode"
        .Add Application.Match("Subtitle", Sheets(seriessheet).Range("A1:Z1"), 0), "subtitle"
    End With
    Set GetSeriesColumnsInSheet = dict
End Function

Public Sub PlusPlus()
    ' Increment the selected cell by one if it is numeric.
    '
    ' If the incremented cell is in the "Ep" column, insert
    ' a date-title-season-episode-subtitle entry in the Episodes sheet
    ' to log the completion of an episode.
    
    If Not CanIncrement() Then
        Exit Sub
    End If
    
    ActiveCell.Value = ActiveCell.Value + 1
    
    Dim logsheet As String
    logsheet = "Episodes"
    Set logCols = GetLogColumnsInSheet(logsheet)
    Set seriesCols = GetSeriesColumnsInSheet(ActiveSheet.Name)
    
    Dim nextrow As Integer
    nextrow = logCols.Item("count").Value + 2
    
    Dim Response
    
    If nextrow < 0 Or nextrow >= 65536 Then
        Response = MsgBox("Error: Did not insert log entry because target row " + CStr(nextrow) + " is out of range.", vbCritical + vbOKOnly, "Row out of range")
        Exit Sub
    End If
        
    ' Set index and increment ref
    
    Dim target As Range
    Dim read As Range
    Set targetcell = Sheets(logsheet).Cells(nextrow, 1) ' output
    Set readcell = ActiveCell.Offset(0, 1 - ActiveCell.Column) ' input
    
    If Not IsEmpty(targetcell.Offset(0, 0)) Then
        Response = MsgBox("Warning: Did not insert log entry because target row " + CStr(nextrow) + " was not empty.", vbCritical + vbOKOnly, "Target not empty")
        Exit Sub
    End If
              
    ' Insert new log row
    
    targetcell.Offset(0, logCols.Item("date") - 1).Formula = "=TEXT(NOW(),""yyyy-mm-dd  HH:mm"")"
    targetcell.Offset(0, logCols.Item("date") - 1).Value = targetcell.Offset(0, logCols.Item("date") - 1).Value
    
    targetcell.Offset(0, logCols.Item("title") - 1).Value = readcell.Offset(0, seriesCols.Item("title") - 1).Value
    targetcell.Offset(0, logCols.Item("episode") - 1).Value = readcell.Offset(0, seriesCols.Item("episode") - 1).Value
    
End Sub
