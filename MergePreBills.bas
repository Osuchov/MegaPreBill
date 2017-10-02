Attribute VB_Name = "MergePreBills"
Option Explicit

Sub Merge()

Dim ws As Worksheet
Dim wsRoad As Worksheet, wsFCL As Worksheet, wsLCL As Worksheet, wsAir As Worksheet, wsALL As Worksheet
Dim wsCons As Worksheet 'consolidated worksheet
Dim file As String
Dim wb As Workbook
Dim wbMerge As Workbook
Dim folder As String
Dim counter As Long, allFiles As Long
Dim target As Range     'where to paste
Dim pb As PreBill
Dim pbNum As Double
Dim fFree As Long, fFreeAll As Long
Dim sht As Variant
Dim arrSheets As Variant
Dim completed As Single

folder = pickDir("Pick the directory with Excel files to merge", "Merge")

If Len(folder) = 0 Then
    MsgBox "Folder not picked. Exiting...", vbExclamation
    Exit Sub
End If

Application.ScreenUpdating = False

Set wbMerge = Application.ThisWorkbook  'naming sheets
Set wsRoad = wbMerge.Sheets("Road")
Set wsFCL = wbMerge.Sheets("FCL")
Set wsLCL = wbMerge.Sheets("LCL")
Set wsAir = wbMerge.Sheets("Air")
Set wsALL = wbMerge.Sheets("ALL")

file = Dir(folder & "*.xls")
allFiles = 0

Do While file <> ""
    allFiles = allFiles + 1
    file = Dir()
Loop

file = Dir(folder & "*.xls")

Do Until Len(file) = 0                  'loop on files to be merged
    counter = counter + 1
    completed = Round((counter * 100) / allFiles, 0)
    progress completed
        
    Set wb = Workbooks.Open(folder & file)
    Set ws = wb.Sheets(1)

    If Range("C1") = "CA11" Then          'Canada new pre bill template
        'Canda's template does not match pb methods - adjustment by adding 1 row
        ws.Rows("9:9").EntireRow.Insert Shift:=xlShiftDown

        If Range("B5") = "" Then
            GoTo Exception          'move on to the next pre bill
        Else
            pbNum = CDbl(Range("B5"))
        End If
    Else                            'For any other company code than CA11
        If Range("B6") = "" Then    'Range("B6") can be empty with volatiles
            GoTo Exception          'move on to the next pre bill
        Else
            pbNum = Range("B6")
        End If
    End If

    Set pb = New PreBill                'setting new pre bill object with attributes from file
    
    pb.CC = Range("C1")                 'setting company code
    pb.Number = pbNum
    pb.CarrierCode = Range("C2")    'setting the rest of pre bill attributes
    pb.Status = Range("B9")
    pb.Vendor = Range("B5")
    pb.Period = Range("B3")
    pb.CreationDate = Range("B7")
    pb.NumberOfColumns = countColumns()
    pb.NumberOfRows = countRows()
    pb.Mode = ws.Name
    
    pb.Copy                         'copying of pre bill atributes
    
    If pb.Mode = "Road" Or pb.Mode = "Road Azkar" Or pb.Mode = "Road US" Then  'determining transport mode (pre bill template)
        fFree = firstFree(wsRoad)                       'checking first free cell in the correct sheet
        Set target = wsRoad.Cells(fFree, 8)             'setting pasting target
    ElseIf pb.Mode = "FCL" Or pb.Mode = "Sea" Then
        fFree = firstFree(wsFCL)
        Set target = wsFCL.Cells(fFree, 8)
    ElseIf pb.Mode = "Air" Or pb.Mode = "Air 2" Then
        fFree = firstFree(wsAir)
        Set target = wsAir.Cells(fFree, 8)
    ElseIf pb.Mode = "Sea LCL" Then
        fFree = firstFree(wsLCL)
        Set target = wsLCL.Cells(firstFree(wsLCL), 8)
    Else
        MsgBox "WARNING! This workbook contains an unknown transport mode name: " & pb.Mode _
        & " Pre bill " & pb.Number & " (" & pb.CarrierCode & "/" & pb.CC & ") will not be merged!"
        counter = counter - 1
        GoTo Exception                                  'unknown transport mode exception
    End If
    
    target.PasteSpecial Paste:=xlPasteValuesAndNumberFormats    'paste the rest of pre bill data
    pb.PastePBData (fFree)                                      'to a first free cell
    
    fFreeAll = firstFree(wsALL)                         'pasting a whole pre bill file to "ALL" Sheet
    If fFreeAll = 2 Then
        Set target = wsALL.Cells(1, 1)
    Else
        Set target = wsALL.Cells(fFreeAll, 1)
    End If
    
    ws.UsedRange.Copy
    target.PasteSpecial Paste:=xlPasteAllExceptBorders
    
Exception:
    Application.CutCopyMode = False
    wb.Close False
    file = Dir
Loop

arrSheets = Array(Road, FCL, LCL, Air, ALL)

For Each sht In arrSheets
    sht.UsedRange.WrapText = False
Next sht

Unload UserForm1
Application.ScreenUpdating = True
MsgBox "Merging of " & counter & " Excel files completed", vbInformation

End Sub

Sub progress(completed As Single)
UserForm1.Text.Caption = completed & "% Completed"
UserForm1.Bar.Width = completed * 2

DoEvents

End Sub

Function pickDir(winTitle As String, buttonTitle As String) As String

Dim window As FileDialog
Dim picked As String

Set window = Application.FileDialog(msoFileDialogFolderPicker)
window.Title = winTitle
window.ButtonName = buttonTitle

If window.Show = -1 Then
    picked = window.SelectedItems(1)
    If Right(picked, 1) <> "\" Then
        pickDir = picked & "\"
    Else
        pickDir = picked
    End If
    
End If

End Function

Function timestamp() As String

timestamp = Format(Now(), "_yyyymmdd_hhmmss")

End Function

Function countColumns() As Long
    countColumns = ActiveSheet.UsedRange.Columns.count
End Function

Function countRows() As Long
    countRows = ActiveSheet.UsedRange.Rows.count - 12
    
End Function

Function firstFree(works As Worksheet) As Long
    works.Activate
    firstFree = ActiveSheet.UsedRange.Rows.count + 1
End Function

Sub clear_all()
Dim mb As Integer
Dim arrSheets As Variant
Dim sht As Variant

mb = MsgBox("You are about to clear all data from pre bill sheets." & Chr(13) & "Are you sure?", vbOKCancel + vbQuestion)

If mb = 1 Then
    arrSheets = Array(Road, FCL, LCL, Air, ALL)
    
    For Each sht In arrSheets
        If sht.Name = "ALL" Then
            sht.Activate
            sht.UsedRange.Select
        Else
            On Error Resume Next   'turn off error reporting
            ActiveSheet.ShowAllData
            sht.Activate
            sht.UsedRange.Offset(1).Select
            On Error GoTo 0
        End If
        Selection.EntireRow.Delete
        Cells(2, 1).Select
    Next sht
    
    MsgBox "Pre bill sheets are now empty."
Else
    MsgBox "Macro cancelled."
End If

End Sub
