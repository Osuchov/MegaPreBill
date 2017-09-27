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
    Cells.Find(what:="HKD018554", After:=ActiveCell, LookIn:=xlValues, _
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

Public Sub ExportSourceFiles()

Dim destPath As String
Dim component As VBComponent

destPath = "C:\Users\PH325251\Desktop\VBA\Git\MegaPreBill\"

For Each component In Application.VBE.ActiveVBProject.VBComponents
If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
component.Export (destPath & component.Name & ToFileExtension(component.Type))
End If
Next
 
End Sub
 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
Select Case vbeComponentType
Case vbext_ComponentType.vbext_ct_ClassModule
ToFileExtension = ".cls"
Case vbext_ComponentType.vbext_ct_StdModule
ToFileExtension = ".bas"
Case vbext_ComponentType.vbext_ct_MSForm
ToFileExtension = ".frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner
Case vbext_ComponentType.vbext_ct_Document
Case Else
ToFileExtension = vbNullString
End Select
 
End Function
