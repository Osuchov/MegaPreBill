Attribute VB_Name = "DisputeCheck"
Option Explicit

Sub CheckDisputes()

Dim fd As Office.FileDialog
Dim wb As Workbook
Dim wsDisputes As Worksheet
Dim disputeFile As String
Dim disputeRng As Range
Dim parkedDispute As Range
Dim Shipment As String
Dim arrSheets As Variant, sht As Variant
Dim lookWhere As Range, foundWhere As Range
Dim firstFoundAddress As String
Dim preBills As String
Dim missingShipmentRow As Long
Dim counter As Long, allDisputes As Long
Dim completed As Single

arrSheets = Array(Road, FCL, LCL, Air)

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
Set wsDisputes = Sheets("Disputes")
wsDisputes.Rows("1:1").AutoFilter Field:=25, Criteria1:="parked", _
    VisibleDropDown:=False

Set disputeRng = wsDisputes.UsedRange.columns(9)

allDisputes = 0

For Each parkedDispute In disputeRng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow
    Shipment = Cells(parkedDispute.row, 9).Value
    If Shipment <> "" Then
        allDisputes = allDisputes + 1
    Else
        Exit For
    End If
Next parkedDispute

counter = 0

'loop on all disputes (shipment numbers)
For Each parkedDispute In disputeRng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow
    counter = counter + 1
    completed = Round((counter * 100) / allDisputes, 0)
    progress completed

    preBills = ""
    Shipment = Cells(parkedDispute.row, 9).Value
    If Shipment = "" Then
        missingShipmentRow = parkedDispute.row
        GoTo missingShipment
    End If
    
    For Each sht In arrSheets   'loop through all transport modes
        Set lookWhere = sht.UsedRange.columns(8)
        Set foundWhere = lookWhere.Find(what:=Shipment, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
        
        If Not foundWhere Is Nothing Then
            firstFoundAddress = foundWhere.Address
            preBills = preBills & sht.Cells(foundWhere.row, 1).Value & " " 'lookWhere.Cells(foundWhere.Row, 1).Value & " "
            Do
                Set foundWhere = lookWhere.FindNext(foundWhere)
                If Not foundWhere Is Nothing Then
                    preBills = preBills & sht.Cells(foundWhere.row, 1).Value & " "
                Else
                    Exit Do
                End If
            Loop While foundWhere.Address <> firstFoundAddress
        End If
    Next sht
    
    If Len(preBills) = 0 Then
        preBills = "not found"
    End If
    Cells(parkedDispute.row, 40).Value = preBills

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
    MsgBox "Shipment in " & missingShipmentRow & " is missing. Check if that is the end of the file."
    Exit Sub
End Sub

Sub progress(completed As Single)
UserForm2.Text.Caption = completed & "% Completed"
UserForm2.Bar.Width = completed * 2

DoEvents

End Sub
