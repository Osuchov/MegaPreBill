Attribute VB_Name = "AdditionalCostsCheck"
Option Explicit

Public Sub CheckAC()
Dim Shipment, shipments As Range
Dim srcShipment As String

AC.Activate
Set shipments = Range(Cells(2, 7), Cells(AC.UsedRange.rows.Count, 7))
shipments.Select

For Each Shipment In shipments
    srcShipment = Shipment.Value
    Cells.Find(what:="HKD018554", After:=ActiveCell, LookIn:=xlValues, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
Next Shipment
    
End Sub
