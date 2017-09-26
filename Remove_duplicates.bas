Attribute VB_Name = "Remove_duplicates"
Sub RemoveDuplicates()

    Dim roadRows As Long, fclRows As Long, lclRows As Long, airRows As Long
    Dim NEWroadRows As Long, NEWfclRows As Long, NEWlclRows As Long, NEWairRows As Long
    
    roadRows = countRows(Road, 1)
    fclRows = countRows(FCL, 1)
    lclRows = countRows(LCL, 1)
    airRows = countRows(Air, 1)
    
    Road.Select
    Columns("A:A").Select
    ActiveSheet.Range("$A$1:$AM$" & roadRows).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5 _
        , 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, _
        33, 34, 35, 36, 37, 38, 39), Header:=xlYes

    FCL.Select
    Columns("A:A").Select
    ActiveSheet.Range("$A$1:$AP$" & fclRows).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, _
        6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33 _
        , 34, 35, 36, 37, 38, 39, 40, 41, 42), Header:=xlYes

    LCL.Select
    Columns("A:A").Select
    ActiveSheet.Range("$A$1:$AQ$" & lclRows).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, _
        6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33 _
        , 34, 35, 36, 37, 38, 39, 40, 41, 42, 43), Header:=xlYes

    Air.Select
    Columns("A:A").Select
    ActiveSheet.Range("$A$1:$AT$" & airRows).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, _
        6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33 _
        , 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46), Header:=xlYes
        
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
    countRows = ws.Cells(1, column).Row
Else
    countRows = ws.Cells(1, column).End(xlDown).Row
End If

End Function
