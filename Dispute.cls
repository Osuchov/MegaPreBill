VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dispute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dShipment As String
Private dPreBill As Long

'Getters
Public Property Get Shipment() As String
    Shipment = dShipment
End Property

Public Property Get PreBill() As Long
    PreBill = dPreBill
End Property

'Setters
Public Property Let Shipment(ship As Long)
    dShipment = ship
End Property

Public Property Let PreBill(pb As Long)
    dShipment = ship
End Property
