Attribute VB_Name = "AC_check"
Option Explicit

Public Sub ACcheck()
Dim shipment, shipments As Range
Dim srcShipment As String

AC.Activate
Set shipments = Range(Cells(2, 7), Cells(AC.UsedRange.Rows.Count, 7))
shipments.Select

For Each shipment In shipments
    srcShipment = shipment.Value
    Cells.Find(What:="HKD018554", After:=ActiveCell, LookIn:=xlValues, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
Next shipment
    
End Sub
