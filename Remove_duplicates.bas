Attribute VB_Name = "Remove_duplicates"
Sub RemoveDuplicates()

    Dim roadRows As Long, fclRows As Long, lclRows As Long, airRows As Long
    Dim NEWroadRows As Long, NEWfclRows As Long, NEWlclRows As Long, NEWairRows As Long
    Dim columns As Long
    Dim arrSheets As Variant, sht As Variant
    Dim colArr()
    
    arrSheets = Array(Road, FCL, LCL, Air)
    
    roadRows = countRows(Road, 1)
    fclRows = countRows(FCL, 1)
    lclRows = countRows(LCL, 1)
    airRows = countRows(Air, 1)
    
    For Each sht In arrSheets
        columns = countCols(Sheets(sht.Name), 1)  'count columns in sht
        ReDim colArr(0 To columns - 1)
        
        For i = 0 To columns - 1
            colArr(i) = i + 1
        Next i
        
        sht.UsedRange.RemoveDuplicates columns:=(colArr), Header:=xlYes
    Next sht
    
    NEWroadRows = countRows(Road, 1)
    NEWfclRows = countRows(FCL, 1)
    NEWlclRows = countRows(LCL, 1)
    NEWairRows = countRows(Air, 1)

MsgBox "Remove duplicates finished." & Chr(13) _
        & "Road duplicates: " & roadRows - NEWroadRows & Chr(13) _
        & "FCL duplicates: " & fclRows - NEWfclRows & Chr(13) _
        & "LCL duplicates: " & lclRows - NEWlclRows & Chr(13) _
        & "Air duplicates: " & airRows - NEWairRows & Chr(13)

End Sub

Function countRows(ws As Worksheet, column As Long) As Long
'finds last used row (with header)

If ws.Cells(2, column) = "" Then
    countRows = ws.Cells(1, column).row
Else
    countRows = ws.Cells(1, column).End(xlDown).row
End If

End Function

Function countCols(ws As Worksheet, row As Long) As Long
'finds last used column in row

If ws.Cells(row, 2) = "" Then
    countCols = ws.Cells(row, 1).column
Else
    countCols = ws.Cells(row, columns.Count).End(xlToLeft).column
End If

End Function
