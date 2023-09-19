Public Sub reauto()
    ' Restore automatic screen updating
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Function inside(ByRef key As Variant, ByRef dict As Collection) As Boolean
    ' Return true if the given key is in the given dict; false otherwise
    Dim value As Variant
    On Error GoTo err
        inside = True
        value = dict(key)
        Exit Function
err:
    inside = False
End Function

Private Function streamdict(stream As String) As Collection
    ' Returns a dictionary of parameters and their values for the STREAM sheet.
    ' Assigns the given STREAM sheet name to the key "self".
    '
    ' Searches cells A1:G1 and assigns the following:
    '    datecol: the column containing "Date"
    '    textcol: the column containing "Description"
    '    fromcol: the column containing "From"
    '    tocol: the column containing "To"
    '    amountcol: the column containing "Amount"
    '
    ' For now, the following cells are hard-coded in:
    '    countref: the cell that counts the number of entries (set to H1)
    '    statusref: if this cell is set to TRUE, Build(False) won't do anything (set to K1)
    Dim dict As Collection
    Set dict = New Collection
    Dim page As String
    With dict
        .Add stream, "self"
        .Add Application.Match("Date", sheets(stream).Range("A1:G1"), 0), "datecol"
        .Add Application.Match("Description", sheets(stream).Range("A1:G1"), 0), "textcol"
        .Add Application.Match("From", sheets(stream).Range("A1:G1"), 0), "fromcol"
        .Add Application.Match("To", sheets(stream).Range("A1:G1"), 0), "tocol"
        .Add Application.Match("Amount", sheets(stream).Range("A1:G1"), 0), "amountcol"
        .Add "H1", "countref"
        .Add "K1", "statusref"
    End With
    Set streamdict = dict
End Function

Private Function pagedict(page As String) As Collection
    ' Returns a dictionary of parameters and their values for the PAGE sheet.
    ' Assigns the PAGE sheet name to the key "self".
    '
    ' Searches cells A1:F1 and assigns the following:
    '    aliascol: the column containing "Alias": this is the short word used for identification in "to" and "from" fields
    '    sheetcol: the column containing "Name": this is the name of the corresponding account Excel sheet
    '    streamcol: the column containing "Stream": this is the column counting the number of entries in STREAM
    '    archivecol: the column containing "Archive": this is the column counting the number of entries in the account sheet
    '
    ' Hard-coded in values:
    '    counterref: the cell containing the number of different accounts (set to I1)
    
    Dim dict As Collection
    Set dict = New Collection
    With dict
        .Add page, "self"
        .Add Application.Match("Alias", sheets(page).Range("A1:F1"), 0), "aliascol"
        .Add Application.Match("Name", sheets(page).Range("A1:F1"), 0), "sheetcol"
        .Add Application.Match("Stream", sheets(page).Range("A1:F1"), 0), "streamcol"
        .Add Application.Match("Archive", sheets(page).Range("A1:F1"), 0), "archivecol"
        .Add "I1", "countref"
    End With
    Set pagedict = dict
End Function

Private Function sheetdict(ByRef pref As Collection) As Collection
    ' Returns a lookup table that maps an alias to the row number
    ' of its entry in the PAGE sheet. Requires an initialized
    ' pagedict reference as a parameter.
    
    Dim lut As Collection   ' Key alias to row number in PAGE
    Dim page As String
    Dim rowoffset As Integer
    rowoffset = 1
    page = pref.Item("self")
    Set lut = New Collection
    
    For i = 1 To sheets(page).Range(pref.Item("countref")).value
        lut.Add i + rowoffset, sheets(page).Cells(i + rowoffset, pref.Item("aliascol")).value
        Next i
    
    Set sheetdict = lut
End Function

Private Function accountdict() As Collection
    ' Returns a dictionary of parameters and their values for a generic account sheet.
    ' Includes the column numbers of certain parameters. For now, they're
    ' hard-coded in. So, the following conditions must be satisfied:
    '    Number of columns: 4
    '    Col 1: the date.
    '    Col 2: the text description.
    '    Col 3: the alias of the "to" / "origin" account
    '    Col 4: the net change in money.
    '
    ' Sign of amounts follow enthalpy convention:
    ' The account represents the system, and 'to' denotes the surroundings.
    ' A negative quantity denotes an exothermic reaction, i.e. money being lost
    ' to the surroundings.
    Dim dict As Collection
    Set dict = New Collection
    With dict
        .Add 4, "numcols"
        .Add 1, "datecol"
        .Add 2, "textcol"
        .Add 3, "tocol"
        .Add 4, "amountcol"
    End With
    Set accountdict = dict
End Function

Public Sub clearsheet(sheet As String, row1 As Variant, col1 As Variant, row2 As Variant, col2 As Variant)
    ' Clears the cells in the range [col1, col2] x [row1, row2] in the specified sheet.
    ' Cell formatting is unaffected.
    If (row2 >= row1 And col2 >= col1) Then             ' Ensure that bounds are valid
        If (sheet = "STREAM" Or sheet = "PAGE") Then    ' STREAM and PAGE should not ever undergo clearing
            MsgBox "Method invocation tried to delete from " & sheet & "!", vbExclamation, "Error"
        Else
            sheets(sheet).Range(sheets(sheet).Cells(row1, col1), sheets(sheet).Cells(row2, col2)).ClearContents
        End If
    End If
End Sub

Public Sub Build(all As Boolean)
    ' Scans the STREAM page and rebuild the sheets that have been modified.
    ' If all is set to true, then all sheets are forcibly rebuilt,
    ' whether they have been updated or not.
    
    ' Reduce lag
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Confirm if all
    If (all) Then
        Dim answer As Integer
        answer = MsgBox("You are changing all accounts. Do you wish to continue?", vbYesNo + vbQuestion)
        If (answer = vbNo) Then
            GoTo finish
        End If
    End If
    
    ' Build the necessary pages.
    Dim sref As Collection, pref As Collection, aref As Collection, aliasrow As Collection, dirty As Collection
    Dim stream As String, page As String, archivecount As String, streamcount As String, alias As String, sheet As String
    Dim counters() As Long
    Dim accounts() As Variant
    Dim thisaccount() As Variant
    Dim fromid As String, toid As String, amount As Variant, time As Variant, text As Variant
    Dim arow As Long, lrow As Long, lcol As Long, urow As Long, ucol As Long
    Dim rowoffset As Integer
    
    rowoffset = 1
    stream = "STREAM"
    page = "PAGE"
    other = "debts"
    Set sref = streamdict(stream)
    Set pref = pagedict(page)
    Set aref = accountdict()
    Set aliasrow = sheetdict(pref)
    Set dirty = New Collection
    amount = sheets(page).Range(pref.Item("countref")).value + rowoffset
    ReDim counters(1 To amount)
    ReDim accounts(1 To amount)
    
    ' Clear pages
    
    For Each r In aliasrow
        archivecount = sheets(page).Cells(r, pref.Item("archivecol")).text
        streamcount = sheets(page).Cells(r, pref.Item("streamcol")).text
        If (archivecount <> streamcount Or all) Then                    ' Update required only if archive count != stream count
            alias = sheets(page).Cells(r, pref.Item("aliascol")).text
            sheet = sheets(page).Cells(r, pref.Item("sheetcol")).text
            dirty.Add sheet, CStr(aliasrow.Item(alias))                          ' dirty maps alias row in PAGE to the name of the sheet
            If (streamcount <> 0) Then
                ReDim thisaccount(1 To streamcount, 1 To aref.Item("numcols")) ' Mark out thisaccount as a temporary cell array
            End If
            accounts(r) = thisaccount
            counters(r) = 1
            If (archivecount <> 0) Then
                clearsheet sheet, 1 + rowoffset, 1, rowoffset + archivecount, 4 ' Clear out the old data
            End If
        End If
        Next r
        
    ' Repopulate pages
    
    For i = 1 To sheets(stream).Range(sref.Item("countref")).value
        fromid = sheets(stream).Cells(rowoffset + i, sref.Item("fromcol")).value
        toid = sheets(stream).Cells(rowoffset + i, sref.Item("tocol")).value
        amount = sheets(stream).Cells(rowoffset + i, sref.Item("amountcol")).value
        time = sheets(stream).Cells(rowoffset + i, sref.Item("datecol")).value
        text = sheets(stream).Cells(rowoffset + i, sref.Item("textcol")).value
                      
        ' Take the row only if the fromid is from a dirty page or if "debts"
        '  is a dirty page and the fromid is not in aliasrow (i.e. "other")
        
        arow = -1
        If (Not inside(fromid, aliasrow)) Then
            If (fromid <> "") Then
                arow = aliasrow.Item(other)
            End If
        Else
            arow = aliasrow.Item(fromid)
        End If
        
        If (inside(CStr(arow), dirty)) Then
            accounts(arow)(counters(arow), aref.Item("datecol")) = time
            accounts(arow)(counters(arow), aref.Item("textcol")) = text
            If (arow = aliasrow.Item(other)) Then
                accounts(arow)(counters(arow), aref.Item("tocol")) = fromid ' hacky: for debts, put target instead of origin
            Else
                accounts(arow)(counters(arow), aref.Item("tocol")) = toid
            End If
            accounts(arow)(counters(arow), aref.Item("amountcol")) = -amount
            counters(arow) = counters(arow) + 1
        End If
        
        ' Take the row only if the toid is from a dirty page or if "debts"
        '  is a dirty page and the toid is not in aliasrow (i.e. "other")
        
        arow = -1
        If (Not inside(toid, aliasrow)) Then
            If (toid <> "") Then
                arow = aliasrow.Item(other)
            End If
        Else
            arow = aliasrow.Item(toid)
        End If
        
        If (inside(CStr(arow), dirty)) Then
            accounts(arow)(counters(arow), aref.Item("datecol")) = time
            accounts(arow)(counters(arow), aref.Item("textcol")) = text
            If (arow = aliasrow.Item(other)) Then
                accounts(arow)(counters(arow), aref.Item("tocol")) = toid ' hacky: for debts, put target instead of origin
            Else
                accounts(arow)(counters(arow), aref.Item("tocol")) = fromid
            End If
            accounts(arow)(counters(arow), aref.Item("amountcol")) = amount
            counters(arow) = counters(arow) + 1
        End If
        
        Next i
        
    ' Write arrays to worksheet
    
    For Each r In aliasrow
        If (inside(CStr(r), dirty)) Then
            sheet = dirty.Item(CStr(r))
            lrow = LBound(accounts(r), 1)
            lcol = LBound(accounts(r), 2)
            urow = UBound(accounts(r), 1)
            ucol = UBound(accounts(r), 2)
            sheets(sheet).Range(sheets(sheet).Cells(rowoffset + lrow, lcol), sheets(sheet).Cells(rowoffset + urow, ucol)) = accounts(r)
        End If
        Next r
    ' Reauto
finish:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub buildsome()
    Build False
End Sub

Public Sub buildall()
    Build True
End Sub
