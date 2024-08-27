VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "ConfigCls"

Private Const totalSheetName As String = "Total"
Public Function getWorkbook() As Workbook
    Set getWorkbook = Application.ThisWorkbook
End Function
Public Function getSheet(sheetName As String) As Worksheet
    Set getSheet = ThisWorkbook.Worksheets(sheetName)
End Function
Public Function getSheet_Total() As Worksheet
    Set getSheet_Total = Me.getSheet(totalSheetName)
End Function
Public Function getTotalSheetName() As String
    getTotalSheetName = totalSheetName
End Function
Public Function getSheet_OT() As Worksheet
    Set getSheet_OT = Me.getSheet("OT")
End Function
Public Function getSheet_Import() As Worksheet
    Set getSheet_Import = Me.getSheet("Import")
End Function
Public Function getSheet_SMS() As Worksheet
    Set getSheet_SMS = Me.getSheet("SMS")
End Function
Public Function getMinimumWage() As String
    getMinimumWage = "10.50"
End Function
Public Function getSecondaryWage() As String
    getSecondaryWage = "14.00"
End Function
Public Function getAllAcountNames() As String()
'includes non accounts because it doesn't matter for the one application
    Dim wb As Workbook: Set wb = getWorkbook()
    Dim ws As Worksheet
    Dim wsNames() As String
    Dim i As Long
    
    ReDim wsNames(wb.Worksheets.count - 1) As String
    
    For i = 0 To UBound(wsNames)
        wsNames(i) = wb.Worksheets(i + 1).Cells(1, 1).value
    Next i
    
    getAllAcountNames = wsNames()
End Function
