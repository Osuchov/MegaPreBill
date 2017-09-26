Attribute VB_Name = "Module1"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("G2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Cells.Find(What:="HKD018554", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Sheets("Additional costs check").Select
    Cells.FindNext(After:=ActiveCell).Activate
    Sheets("FCL").Select
    Cells.FindNext(After:=ActiveCell).Activate
End Sub
