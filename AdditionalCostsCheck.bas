Attribute VB_Name = "AdditionalCostsCheck"
Option Explicit

Sub CheckAdditionalCosts()

Dim fd As Office.FileDialog
Dim lastMappingRow As Long
Dim translation As range
Dim wb As Workbook
Dim wsACs As Worksheet
Dim ACRng As range
Dim ACFile As String
Dim parkedAC As range
Dim shipment, carrier As String
Dim missingShipmentRow As Long
Dim counter As Long, allACs As Long
Dim completed As Single
Dim strPreBills As String
Dim GeneralCN As String

ThisWorkbook.Worksheets("Mapping").Activate

With ThisWorkbook.Worksheets("Mapping")
    lastMappingRow = .UsedRange.Rows.Count
    Set translation = .range(.Cells(1, 1), .Cells(lastMappingRow, 3))
End With

Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    .AllowMultiSelect = False
    .Title = "Please select the additional costs file."     'Set the title of the dialog box.
    .Filters.Clear                                          'Clear out the current filters, and add our own.
    .Filters.Add "All Files", "*.*"
    If .Show = True Then                                    'show the dialog box. Returns true if >=1 file picked
        ACFile = .SelectedItems(1)                     'replace txtFileName with your textbox
    End If
End With

On Error GoTo ErrHandling       'turning off warnings
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
Application.ScreenUpdating = False

Set wb = Workbooks.Open(ACFile)

If wb.Sheets(1).Name <> "Additional Costs" Then     'checks if 1st sheet is called "Additional Costs"
    MsgBox "This is not a valid additional costs file. Check it and run macro again."
    GoTo CleaningUp
End If

Set wsACs = Sheets("Additional Costs")
wsACs.Activate
wsACs.Rows("1:1").AutoFilter Field:=29, Criteria1:="Parked"     'filters parked Additional Costs

Set ACRng = wsACs.UsedRange.columns(7)                  'shipment number is in column 7

allACs = 0                                     'additional costs counter

'loop counts how many additional costs are parked
For Each parkedAC In ACRng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow
    shipment = Cells(parkedAC.row, 7).Value
    If shipment <> "" Then
        allACs = allACs + 1
    Else
        Exit For
    End If
Next parkedAC

counter = 0         'loop counter

'loop on all additional costs (shipment numbers)
For Each parkedAC In ACRng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow
    counter = counter + 1
    completed = Round((counter * 100) / allACs, 0)              'for progress bar
    progress completed

    shipment = Trim(Cells(parkedAC.row, 7).Value)          'get shipment number for find function

    If shipment = "" Then
        missingShipmentRow = parkedAC.row
        GoTo missingShipment
    End If

    carrier = Trim(Cells(parkedAC.row, 5).Value)           'get carrier name of the dispute
    carrier = GeneralCarrierName(carrier, translation)          'get general carrier name

    strPreBills = findPreBillForACShipment(shipment, carrier)
    On Error GoTo ErrHandling

    Cells(parkedAC.row, 36).Value = strPreBills            'write found pre bills in Excel

Next parkedAC


CleaningUp:
    On Error Resume Next
    Application.DisplayAlerts = True            'turning on warnings
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
    Unload UserForm3
    Exit Sub

ErrHandling:
    MsgBox Err & ". " & Err.Description
    Resume CleaningUp
missingShipment:
    On Error Resume Next
    Application.DisplayAlerts = True            'turning on warnings
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True

    Unload UserForm3
    MsgBox "Shipment in row " & missingShipmentRow & " is missing. Check if that is the end of the file."
    Exit Sub
End Sub

Sub progress(completed As Single)
UserForm3.Text.Caption = completed & "% Completed"
UserForm3.Bar.Width = completed * 2

DoEvents

End Sub

Function findPreBillForACShipment(shpmnt As Variant, crrr As String) As String
'returns a string of pre bill numbers where searched shipment-carrier pair was found

Dim preBills()
Dim arrSheets As Variant, sht As Variant
Dim lookWhere As range, foundwhere As range
Dim PBCrrr As Variant   'pre bill carrier name
Dim firstFoundAddress As String
Dim check As Integer, check2 As Integer
Dim strPreBills As String
Dim uniquePreBills As New Collection, a
Dim i As Integer
Dim shee As String
Dim mergeFile As String

mergeFile = ThisWorkbook.Name
arrSheets = Array(Road, RoadUS, FCL, LCL, Air, Air2)

ReDim preBills(0 To 0)      'resetting found preBills array
strPreBills = ""            'found pre bill collection to array

For Each sht In arrSheets   'loop through all transport modes
    shee = sht.Name
    Set lookWhere = sht.UsedRange.columns(9)    'shipment number is in column 9
    Set foundwhere = lookWhere.Find(what:=shpmnt, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False)

    If Not foundwhere Is Nothing Then   'if found
        firstFoundAddress = foundwhere.Address  'remember first found address
        PBCrrr = foundwhere.Offset(0, -5).Value
        check = InStr(LCase(PBCrrr), LCase(crrr))
        check2 = checkIfActivityCodeCellIsEmpty(foundwhere, mergeFile, shee)

        If check <> 0 And check2 <> 0 Then
            preBills(UBound(preBills)) = sht.Cells(foundwhere.row, 1).Value     'allocate first found element
        End If

        Do  'loop for FindNext until found address = first found address
            Set foundwhere = lookWhere.FindNext(foundwhere)

            If Not foundwhere Is Nothing Then    'if found again with correct carrier
                PBCrrr = foundwhere.Offset(0, -5).Value
                check = InStr(LCase(PBCrrr), LCase(crrr))
                check2 = checkIfActivityCodeCellIsEmpty(foundwhere, mergeFile, shee)

                If check <> 0 And check2 <> 0 Then
                    ReDim Preserve preBills(0 To UBound(preBills) + 1)              'allocate next found element
                    preBills(UBound(preBills)) = sht.Cells(foundwhere.row, 1).Value 'assign it to the array
                End If
            Else
                Exit Do                         'if not found again
            End If

        Loop While foundwhere.Address <> firstFoundAddress
    End If
Next sht

On Error Resume Next

For Each a In preBills      'creating unique pre bill collection
        If Not IsEmpty(a) Then
            uniquePreBills.Add a, Str(a)
        End If
    Next a

If uniquePreBills.Count = 0 Then
    strPreBills = "not found"               'if collection is empty
ElseIf uniquePreBills.Count = 1 Then
    strPreBills = uniquePreBills.Item(1)    'if there's 1 item in collection
Else
    For i = 1 To uniquePreBills.Count
        strPreBills = (CStr(uniquePreBills.Item(i))) & " " & strPreBills    'if there's more items
    Next i
End If

Set uniquePreBills = Nothing

findPreBillForACShipment = strPreBills

End Function

Public Function checkIfActivityCodeCellIsEmpty(cellRange As range, file As String, sheet As String) As Integer

Dim ActivityCodeColumn As Long
Dim ColumnNumber As Long
Dim testCell As range
Dim testCellValue As String
Dim row As Long
Dim ws As Worksheet

Set ws = Workbooks(file).Worksheets(sheet)
row = cellRange.row

ColumnNumber = findColumnNum("Activity code", file, sheet)
Set testCell = ws.Cells(row, ColumnNumber)
testCellValue = testCell.Value

If testCellValue = "" Then
    checkIfActivityCodeCellIsEmpty = 1
Else
    checkIfActivityCodeCellIsEmpty = 0
End If


End Function

Function findColumnNum(columnName As String, file As String, sheet As String) As Long
'searches for a given column name in given sheet (row 1) and returns its number

Dim searchRange As range
Dim cell As range
Dim col As Long
Dim komorka As String
Dim ws As Worksheet

Set ws = Workbooks(file).Worksheets(sheet)
Set searchRange = ws.range(ws.Cells(1, 1), ws.Cells(1, 50))
col = 0

For Each cell In searchRange
    komorka = cell.Value
    If cell.Value = columnName Then
        col = cell.column
        Exit For
    End If
Next cell

findColumnNum = col

End Function

Sub test()

Dim ColumnNumber As Long
Dim sheet As String


sheet = "Road"
ColumnNumber = findColumnNum("Activity code", sheet)

End Sub
