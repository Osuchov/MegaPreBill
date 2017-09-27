Attribute VB_Name = "Disputes"
Option Explicit

Public Sub checkDisputes()

Dim fd As Office.FileDialog
Dim wb As Workbook
Dim wsDisputes As Worksheet
Dim disputeFile As String
Dim disputeRng As Range
Dim parkedDispute As Range
Dim shipment As String
Dim arrSheets As Variant, sht As Variant
Dim lookWhere As Range
Dim foundWhere As Range

arrSheets = Array(Road, FCL, LCL, Air)

Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    .AllowMultiSelect = False
    .Title = "Please select the dispute file."  'Set the title of the dialog box.
    .Filters.Clear                            'Clear out the current filters, and add our own.
    .Filters.Add "Excel 2003", "*.xls"
    .Filters.Add "Excel 2003", "*.xlsb"
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

Set disputeRng = wsDisputes.UsedRange

For Each parkedDispute In disputeRng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow
    
    shipment = Cells(parkedDispute.Row, 9)
      
    For Each sht In arrSheets
        Set lookWhere = sht.UsedRange.Columns(9)
        Set foundWhere = lookWhere.Find(what:=shipment, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not foundWhere Is Nothing Then
            Cells(parkedDispute.Row, 40) = "Found on pre bill: " & sht.Cells(foundWhere.Row, 1)
            Exit For
        End If
    Next sht

Next parkedDispute


CleaningUp:
    On Error Resume Next
    Application.DisplayAlerts = True 'turning on warnings
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
    Exit Sub
 
ErrHandling:
    MsgBox Err & ". " & Err.Description
    Resume CleaningUp
    
End Sub
