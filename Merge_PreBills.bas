Attribute VB_Name = "Merge_PreBills"
Option Explicit

Sub Merge()

Dim ws As Worksheet
Dim wsRoad As Worksheet, wsFCL As Worksheet, wsLCL As Worksheet, wsAir As Worksheet
Dim wsCons As Worksheet 'consolidated worksheet
Dim file As String
Dim wb As Workbook
Dim wbMerge As Workbook
Dim folder As String
Dim counter As Long
Dim target As Range     'where to paste
Dim pb As PreBill
Dim pbNum As Long
Dim fFree As Long

folder = pickDir("Pick the directory with Excel files to merge", "Merge")

If Len(folder) = 0 Then
    MsgBox "Folder not picked. Exiting...", vbExclamation
    Exit Sub
End If

Application.ScreenUpdating = False

file = Dir(folder & "*.xls")

Set wbMerge = Application.Workbooks("Merge PreBills.xlsb")
Set wsRoad = wbMerge.Sheets("Road")
Set wsFCL = wbMerge.Sheets("FCL")
Set wsLCL = wbMerge.Sheets("LCL")
Set wsAir = wbMerge.Sheets("Air")

Do Until Len(file) = 0
    counter = counter + 1
    Application.StatusBar = "Merging file number: " & counter
    Set wb = Workbooks.Open(folder & file)
    Set ws = wb.Sheets(1)

    Set pb = New PreBill
    
    If Range("B6") = "" Then    'Range("B6") can be empty with volatiles
        pbNum = 0               'assign 0 value to a pre bill number then
    Else
        pbNum = Range("B6")
    End If
    
    pb.Number = pbNum
    pb.CC = Range("C1")
    pb.CarrierCode = Range("C2")
    pb.Status = Range("B9")
    pb.Vendor = Range("B5")
    pb.Period = Range("B3")
    pb.CreationDate = Range("B7")
    pb.NumberOfColumns = countColumns()
    pb.NumberOfRows = countRows()
    pb.Mode = ws.Name
    
    pb.Copy
    
    If pb.Mode = "Road" Or pb.Mode = "Road Azkar" Then
        fFree = firstFree(wsRoad)
        Set target = wsRoad.Cells(fFree, 8)
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
        GoTo Exception
    End If
    
    target.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    pb.PastePBData (fFree)
    
Exception:
    Application.CutCopyMode = False
    wb.Close False
    file = Dir
Loop

Application.ScreenUpdating = True
Application.StatusBar = False
MsgBox "Merging of " & counter & " Excel files completed", vbInformation

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
    countColumns = ActiveSheet.UsedRange.Columns.Count
End Function

Function countRows() As Long
    countRows = ActiveSheet.UsedRange.Rows.Count - 12
    
End Function

Function firstFree(works As Worksheet) As Long
    works.Activate
    firstFree = ActiveSheet.UsedRange.Rows.Count + 1
End Function

Sub clear_all()
Dim mb As Integer
Dim arrSheets As Variant
Dim sht As Variant

mb = MsgBox("You are about to clear all data from pre bill sheets." & Chr(13) & "Are you sure?", vbOKCancel + vbQuestion)

If mb = 1 Then
    arrSheets = Array(Road, FCL, LCL, Air)
    
    For Each sht In arrSheets
        sht.UsedRange.Offset(1).ClearContents
    Next sht
    
    MsgBox "Pre bill sheets are now empty."
Else
    MsgBox "Macro cancelled."
End If

End Sub
