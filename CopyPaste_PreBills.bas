Attribute VB_Name = "CopyPaste_PreBills"
Option Explicit

Public Sub CopyPaste_PreBills()
Dim folder As String
Dim file As String
Dim counter As Long
Dim fFree As Long
Dim target As Range     'where to paste
Dim wsALL As Worksheet
Dim ws As Worksheet

Dim wb As Workbook
Dim wbMerge As Workbook

folder = pickDir("Pick the directory with Excel files to copy", "Copy")

If Len(folder) = 0 Then
    MsgBox "Folder not picked. Exiting...", vbExclamation
    Exit Sub
End If

'Application.ScreenUpdating = False

file = Dir(folder & "*.xls")

Set wbMerge = Application.ThisWorkbook
Set wsALL = wbMerge.Sheets("ALL")

Do Until Len(file) = 0
    counter = counter + 1
    Application.StatusBar = "Copying file number: " & counter
    Set wb = Workbooks.Open(folder & file)
    Set ws = wb.Sheets(1)
    
    fFree = firstFree(wsALL)
    If fFree = 2 Then
        Set target = wsALL.Cells(1, 1)
    Else
        Set target = wsALL.Cells(fFree, 1)
    End If
    
    ws.UsedRange.Copy
    target.PasteSpecial Paste:=xlPasteAllExceptBorders
    
    Application.CutCopyMode = False
    wb.Close False
    file = Dir
Loop

wsALL.UsedRange.WrapText = False
Application.ScreenUpdating = True
Application.StatusBar = False

MsgBox "Coping of " & counter & " Excel files completed", vbInformation
End Sub
