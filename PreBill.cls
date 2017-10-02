VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PreBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pNumber As Double
Private pCc As String
Private pCarrierCode As String
Private pStatus As String
Private pVendor As String
Private pPeriod As String
Private pCreationDate As String
Private pNumberOfColumns As Long
Private pNumberOfRows As Long
Private pMode As String

'Getters
Public Property Get Number() As Double
    Number = pNumber
End Property
Public Property Get CC() As String
    CC = pCc
End Property
Public Property Get CarrierCode() As String
    CarrierCode = pCarrierCode
End Property
Public Property Get Status() As String
    Status = pStatus
End Property
Public Property Get Vendor() As String
    Vendor = pVendor
End Property
Public Property Get Period() As String
    Period = pPeriod
End Property
Public Property Get CreationDate() As String
    CreationDate = pCreationDate
End Property
Public Property Get NumberOfColumns() As Long
    NumberOfColumns = pNumberOfColumns
End Property
Public Property Get NumberOfRows() As Long
    NumberOfRows = pNumberOfRows
End Property
Public Property Get Mode() As String
    Mode = pMode
End Property

'Setters
Public Property Let Number(num As Double)
    pNumber = num
End Property
Public Property Let CC(comp_code As String)
    pCc = comp_code
End Property
Public Property Let CarrierCode(CCode As String)
    pCarrierCode = CCode
End Property
Public Property Let Status(stat As String)
    pStatus = stat
End Property
Public Property Let Vendor(vend As String)
    pVendor = vend
End Property
Public Property Let Period(per As String)
    pPeriod = per
End Property
Public Property Let CreationDate(crDate As String)
    pCreationDate = crDate
End Property
Public Property Let NumberOfColumns(colNum As Long)
    pNumberOfColumns = colNum
End Property
Public Property Let NumberOfRows(rowNum As Long)
    pNumberOfRows = rowNum
End Property
Public Property Let Mode(transMode As String)
    pMode = transMode
End Property

Public Function Copy()
    Set pbBody = Range(Cells(12, 1), Cells((NumberOfRows + 11), NumberOfColumns))
    pbBody.Select
    Selection.Copy
End Function

Public Function PastePBData(firstFreeRow As Long)
    Range(Cells(firstFreeRow, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 1)).Value = Number()
    Range(Cells(firstFreeRow, 2), Cells(ActiveSheet.UsedRange.Rows.Count, 2)).Value = CC()
    Range(Cells(firstFreeRow, 3), Cells(ActiveSheet.UsedRange.Rows.Count, 3)).Value = CarrierCode()
    Range(Cells(firstFreeRow, 4), Cells(ActiveSheet.UsedRange.Rows.Count, 4)).Value = Status()
    Range(Cells(firstFreeRow, 5), Cells(ActiveSheet.UsedRange.Rows.Count, 5)).Value = Vendor()
    Range(Cells(firstFreeRow, 6), Cells(ActiveSheet.UsedRange.Rows.Count, 6)).Value = Period()
    Range(Cells(firstFreeRow, 7), Cells(ActiveSheet.UsedRange.Rows.Count, 7)).Value = CreationDate()
End Function
