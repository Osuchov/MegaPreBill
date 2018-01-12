Attribute VB_Name = "DisputeCheck"
Option Explicit

Sub CheckDisputes()

Dim fd As Office.FileDialog
Dim wb As Workbook
Dim wsDisputes As Worksheet
Dim disputeFile As String
Dim disputeRng As Range
Dim parkedDispute As Range
Dim shipment, carrier As String
Dim translation As Range
Dim lastMappingRow As Long
Dim missingShipmentRow As Long
Dim counter As Long, allDisputes As Long
Dim completed As Single
Dim strPreBills As String
Dim GeneralCN As String

ThisWorkbook.Worksheets("Mapping").Activate

With ThisWorkbook.Worksheets("Mapping")
    lastMappingRow = .UsedRange.rows.Count
    Set translation = .Range(.Cells(1, 1), .Cells(lastMappingRow, 3))
End With

Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    .AllowMultiSelect = False
    .Title = "Please select the dispute file."  'Set the title of the dialog box.
    .Filters.Clear                            'Clear out the current filters, and add our own.
    .Filters.Add "All Files", "*.*"
    If .Show = True Then                'show the dialog box. Returns true if >=1 file picked
        disputeFile = .SelectedItems(1) 'replace txtFileName with your textbox
    End If
End With

On Error GoTo ErrHandling       'turning off warnings
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
Application.ScreenUpdating = False

Set wb = Workbooks.Open(disputeFile)

If wb.Sheets(1).Name <> "Disputes" Then     'checks if 1st sheet is called "Disputes"
    MsgBox "This is not a valid dispute file. Check it and run macro again."
    GoTo CleaningUp
End If

Set wsDisputes = Sheets("Disputes")
wsDisputes.rows("1:1").AutoFilter Field:=25, Criteria1:="parked" 'filters parked disputes

Set disputeRng = wsDisputes.UsedRange.columns(9)    'shipment number is in column 9

allDisputes = 0                                     'dispute counter

'loop counts how many disputes are parked
For Each parkedDispute In disputeRng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow
    shipment = Cells(parkedDispute.row, 9).Value
    If shipment <> "" Then
        allDisputes = allDisputes + 1
    Else
        Exit For
    End If
Next parkedDispute

counter = 0         'loop counter

'loop on all disputes (shipment numbers)
For Each parkedDispute In disputeRng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow
    counter = counter + 1
    completed = Round((counter * 100) / allDisputes, 0) 'for progress bar
    progress completed

    shipment = Trim(Cells(parkedDispute.row, 9).Value)          'get shipment number for find function
    
    If shipment = "" Then
        missingShipmentRow = parkedDispute.row
        GoTo missingShipment
    End If
    
    carrier = Trim(Cells(parkedDispute.row, 8).Value)           'get carrier name of the dispute
    carrier = GeneralCarrierName(carrier, translation)          'get general carrier name
    
    strPreBills = findPreBillForShipment(shipment, carrier)
    On Error GoTo ErrHandling
    
    Cells(parkedDispute.row, 36).Value = strPreBills            'write found pre bills in Excel
    
Next parkedDispute


CleaningUp:
    On Error Resume Next
    Application.DisplayAlerts = True 'turning on warnings
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
    Unload UserForm2
    Exit Sub
 
ErrHandling:
    MsgBox Err & ". " & Err.Description
    Resume CleaningUp
missingShipment:
    On Error Resume Next
    Application.DisplayAlerts = True 'turning on warnings
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
    
    Unload UserForm2
    MsgBox "Shipment in row " & missingShipmentRow & " is missing. Check if that is the end of the file."
    Exit Sub
End Sub

Sub progress(completed As Single)
UserForm2.Text.Caption = completed & "% Completed"
UserForm2.Bar.Width = completed * 2

DoEvents

End Sub

Function GeneralCarrierName(FullCarrierName As String, translation As Range) As String 'returns a general carrier name from full carrier name

On Error GoTo NotFound
GeneralCarrierName = Application.WorksheetFunction.VLookup(FullCarrierName, translation, 3, 0)
On Error GoTo 0

Exit Function

NotFound:
GeneralCarrierName = FullCarrierName

End Function

Function findPreBillForShipment(shpmnt As Variant, crrr As String) As String
'returns a string of pre bill numbers where searched shipment-carrier pair was found

Dim preBills()
Dim arrSheets As Variant, sht As Variant
Dim lookWhere As Range, foundWhere As Range
Dim PBCrrr As Variant   'pre bill carrier name
Dim firstFoundAddress As String
Dim check As Integer
Dim strPreBills As String
Dim uniquePreBills As New Collection, a
Dim i As Integer
Dim shee As String

arrSheets = Array(Road, RoadUS, FCL, LCL, Air)

ReDim preBills(0 To 0)      'resetting found preBills array
strPreBills = ""            'found pre bill collection to array
    
For Each sht In arrSheets   'loop through all transport modes
    shee = sht.Name
    Set lookWhere = sht.UsedRange.columns(9)    'shipment number is in column 9
    Set foundWhere = lookWhere.Find(what:=shpmnt, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False)
            
    If Not foundWhere Is Nothing Then   'if found
        firstFoundAddress = foundWhere.Address  'remember first found address
        PBCrrr = foundWhere.Offset(0, -5).Value
        check = InStr(LCase(PBCrrr), LCase(crrr))
        
        If check <> 0 Then
            preBills(UBound(preBills)) = sht.Cells(foundWhere.row, 1).Value     'allocate first found element
        End If
        
        Do  'loop for FindNext until found address = first found address
            Set foundWhere = lookWhere.FindNext(foundWhere)

            If Not foundWhere Is Nothing Then    'if found again with correct carrier
                PBCrrr = foundWhere.Offset(0, -5).Value
                check = InStr(LCase(PBCrrr), LCase(crrr))
                
                If check <> 0 Then
                    ReDim Preserve preBills(0 To UBound(preBills) + 1)              'allocate next found element
                    preBills(UBound(preBills)) = sht.Cells(foundWhere.row, 1).Value 'assign it to the array
                End If
            Else
                Exit Do                         'if not found again
            End If
            
        Loop While foundWhere.Address <> firstFoundAddress
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

findPreBillForShipment = strPreBills

End Function
