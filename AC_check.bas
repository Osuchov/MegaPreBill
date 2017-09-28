Attribute VB_Name = "AC_check"
Option Explicit

Public Sub ACcheck()
Dim Shipment, shipments As Range
Dim srcShipment As String

AC.Activate
Set shipments = Range(Cells(2, 7), Cells(AC.UsedRange.Rows.count, 7))
shipments.Select

For Each Shipment In shipments
    srcShipment = Shipment.Value
    Cells.Find(what:="HKD018554", After:=ActiveCell, LookIn:=xlValues, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
Next Shipment
    
End Sub
