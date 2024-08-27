VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangesCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "RangesCls"

Private Const regWageName As String = "RegWage"
Private Const sortFieldName As String = "SortField"
Private Const empName As String = "Emp"
Private Const dataName As String = "Data"

Public Function getPasswordRange() As Range
    Set getPasswordRange = Range("Password")
End Function
Public Function getRegWageRange() As Range
    Set getRegWageRange = Range(regWageName)
End Function
Public Function getPayPeriodRange() As Range
    Set getPayPeriodRange = Range("PayPeriod")
End Function
Public Function getPayDayRange(oneToSixteen As Long) As Range
    Set getPayDayRange = Range("PayDay" & oneToSixteen)
End Function
Public Function getHoursSpentTrackerRange() As Range
    Set getHoursSpentTrackerRange = Range("HoursSpentTracker")
End Function
Public Function getTotalEmpRange() As Range
    Set getTotalEmpRange = Me.getEmpRange(Config.getTotalSheetName())
End Function
Public Function getTotalMiscRange() As Range
    Set getTotalMiscRange = Range(Config.getTotalSheetName() & "Misc")
End Function
Public Function getTotalRegWageRange() As Range
    Set getTotalRegWageRange = Range(Config.getTotalSheetName() & regWageName)
End Function
Public Function getTotalSortFieldRange() As Range
    Set getTotalSortFieldRange = Range(Config.getTotalSheetName() & sortFieldName)
End Function
Public Function getTotalNonHrlyRange() As Range
    Set getTotalNonHrlyRange = Range(Config.getTotalSheetName() & "NonHrly")
End Function
Public Function getTotalSecWageRange() As Range
    Set getTotalSecWageRange = Range(Config.getTotalSheetName() & "SecWage")
End Function
Public Function getTotalTotalsRange() As Range
    Set getTotalTotalsRange = Range(Config.getTotalSheetName() & "Totals")
End Function
Public Function getTotalVarianceRange() As Range
    Set getTotalVarianceRange = Range(Config.getTotalSheetName() & "Variance")
End Function
Public Function getEmpRange(sheetName As String) As Range
    Set getEmpRange = Range(sheetName & empName)
End Function
Public Function getSortFieldRange(sheetName As String) As Range
    Set getSortFieldRange = Range(sheetName & sortFieldName)
End Function
Public Function getDataRange(sheetName As String) As Range
    Set getDataRange = Range(sheetName & dataName)
End Function
Public Function getData1Range(sheetName As String) As Range
    Set getData1Range = Range(sheetName & dataName & 1)
End Function
Public Function getData2Range(sheetName As String) As Range
    Set getData2Range = Range(sheetName & dataName & 2)
End Function
Public Function getCopyFromRange(sheetName As String) As Range
    Set getCopyFromRange = Range(sheetName & "CopyFrom")
End Function
Public Function getPasteToRange(sheetName As String) As Range
    Set getPasteToRange = Range(sheetName & "PasteTo")
End Function
Public Function getEmployeeListRange() As Range
    Set getEmployeeListRange = Range("EmployeeList")
End Function
Public Function getImportTableRange() As Range
    Set getImportTableRange = Range("ImportTable")
End Function
Public Function getManagerRange() As Range
    Set getManagerRange = Range("ManagerRange")
End Function
Public Function getHumanityPasswordRange() As Range
    Set getHumanityPasswordRange = Range("HumanityPassword")
End Function
Public Function getOTEmpRange() As Range
    Set getOTEmpRange = getEmpRange(Config.getSheet_OT().name)
End Function
Public Function getHumanityImportsRange() As Range

    Dim im As Worksheet
    Dim a1 As Range
    Dim lRow As Long
    Dim lCol As Long
    Dim sRange As Range
    Dim eRange As Range
    Dim fullRange As Range
    
    Set im = Config.getSheet_Import()
    Set a1 = im.Range("A1")
    
    If a1.Offset(1).value = vNil Or isEmpty(a1.Offset(1)) Then
        MsgBox "There is no imported data", vbInformation, "Nope!"
        End
    End If
    
    lCol = a1.End(xlToRight).Column
    lRow = a1.End(xlDown).Row
    
    
    Set sRange = im.Range(a1.Offset(1), a1.Offset(1, lCol - 1))
    Set eRange = im.Range(a1.Offset(lRow - 1), a1.Offset(lRow - 1, lCol - 1))
    
    Set fullRange = im.Range(sRange, eRange)
    
    Set getHumanityImportsRange = fullRange
    
End Function
Public Function getRangeNames(Optional sheetName As String = vNil) As String()
    
    Dim rangeNames As Names
    Dim rNames() As String
    Dim tempNames() As String
    Dim r As Long
    Dim c As Collection
    
    Set rangeNames = Config.getWorkbook().Names
    ReDim tempNames(rangeNames.count - 1) As String
    
    For r = 1 To rangeNames.count
    
        tempNames(r - 1) = rangeNames(r).name
    
    Next r
    
    If sheetName = vNil Then
        rNames() = tempNames()
    Else
        Set c = New Collection
        
        For r = 0 To UBound(tempNames())
            
            If Words.startsWith(tempNames(r), sheetName) Then
                
                Call c.add(tempNames(r))
                
            End If
            
        Next r
        
        rNames() = Collections.toArray(c)
        
    End If
    
    getRangeNames = rNames()
    
End Function
Public Function getEmpRanges() As String()

    Dim rangeNames() As String
    Dim r As Names
    Dim ns() As String
    Dim n As Long
    Dim c As New Collection
    
    rangeNames() = Ranges.getRangeNames()
    
    For n = 0 To UBound(rangeNames())
    
        If Words.endsWith(rangeNames(n), "Emp") Then
        
            c.add rangeNames(n)
        
        End If
        
    Next n
    
    ns() = Collections.toArray(c)
    
    getEmpRanges = ns()
    
End Function
