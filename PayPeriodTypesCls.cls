VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PayPeriodTypesCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "PayPeriodTypesCls"

Private Const total As String = "T"
Private Const semimonthly As String = "S"
Private Const monthly As String = "M"
Private Const invoice As String = "i"
Private Const overTime As String = "OT"

Public Function getSheetType(sheetName As String) As String

    Dim commentText As String
    Dim sheetType As String
    
    commentText = Config.getSheet(sheetName).Range("A1").Comment.text
    
    sheetType = Strings.Split(commentText, ";")(0)
    
    getSheetType = Strings.Replace(Strings.Trim$(sheetType), vbNewLine, vNil)

End Function
Public Function isTotalSheet(sheetName As String) As Boolean
    isTotalSheet = Me.isTotalType(Me.getSheetType(sheetName))
End Function
Public Function isTotalType(sheetType As String) As Boolean
    isTotalType = sheetType = total
End Function
Public Function isSemimonthlySheet(sheetName As String) As Boolean
    isSemimonthlySheet = Me.isSemimonthlyType(Me.getSheetType(sheetName))
End Function
Public Function isSemimonthlyType(sheetType As String) As Boolean
    isSemimonthlyType = sheetType = semimonthly
End Function
Public Function isMonthlySheet(sheetName As String) As Boolean
    isMonthlySheet = Me.isMonthlyType(Me.getSheetType(sheetName))
End Function
Public Function isMonthlyType(sheetType As String) As Boolean
    isMonthlyType = sheetType = monthly
End Function
Public Function isInvoiceSheet(sheetName As String) As Boolean
    isInvoiceSheet = Me.isInvoiceSheet(Me.getSheetType(sheetName))
End Function
Public Function isInvoiceType(sheetType As String) As Boolean
    isInvoiceType = Strings.Right$(sheetType, 1) = invoice
End Function
Public Function isOvertimeSheet(sheetName As String) As Boolean
    isOvertimeSheet = Me.isOvertimeType(Me.getSheetType(sheetName))
End Function
Public Function isOvertimeType(sheetType As String) As Boolean
    isOvertimeType = sheetType = overTime
End Function
