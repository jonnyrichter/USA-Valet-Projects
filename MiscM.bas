Attribute VB_Name = "MiscM"
Option Explicit

'Clears all import data and removes all employee names from all sheets.
Public Sub resetData()
    
    Dim n() As String
    Dim i As Long
    Dim misc As Range
    Dim regWage As Range
    Dim secWage As Range
    Dim emp As Range
    Dim importRange As Range
    
    Set misc = Ranges.getTotalMiscRange()
    Set regWage = Ranges.getTotalRegWageRange()
    Set secWage = Ranges.getTotalSecWageRange()
    Set importRange = Ranges.getHumanityImportsRange()
    
    n() = Ranges.getEmpRanges()
    
    System.Update False
    
    Call UnprotectAllSheetsm.UnprotectAllSheets
    
    misc.ClearContents
    misc.ClearComments
    regWage.value = Config.getMinimumWage()
    secWage.value = Config.getSecondaryWage()
    
    importRange.ClearContents
    importRange.ClearComments
    importRange.Interior.Color = xlNone
    
    For i = 0 To UBound(n())
    
        Set emp = Range(n(i))
        
        emp.ClearComments
        emp.ClearContents
    
    Next i
    
    Call UnprotectAllSheetsm.ProtectAllSheets
    
    System.Update True

End Sub

'I have no idea why I made this
Public Sub fillOTColumnNameReferences()
    fillColumns "BG", "BZ"
End Sub
Private Sub fillColumns(startColumn As String, endColumn As String)
    Dim ot As Worksheet
    Dim startRange As Range, endRange As Range, totalRange As Range
    Dim c As Range
    Dim rowStart As Long
    
    Set ot = Config.getSheet_OT()
    Set startRange = ot.Range(startColumn & "1")
    Set endRange = ot.Range(endColumn & "1")
    Set totalRange = Range(startRange, endRange)
    
    Set c = startRange.Offset(0, -1)
    
    rowStart = CLng(Replace(c.Formula, "=Total!$AH", ""))
    
    For Each c In totalRange
        c.Formula = "=Total!$AH" & rowStart + 1
        rowStart = rowStart + 1
    Next c

End Sub
