Attribute VB_Name = "NewPayPeriodM"
Option Explicit

Private Const getClass As String = "NewPayPeriodM"

Sub NewPayPeriod()

Dim ws As Worksheet, WsType As String, WsName As String
Dim dataRange As Range, dataRange1 As Range, dataRange2 As Range
Dim copyFromRange As Range, pasteToRange As Range
Dim payPeriod As Range, payDay1 As Range, payDay13 As Range, payDay14 As Range, payDay15 As Range, payDay16 As Range
Dim Cell As Variant
Dim FYear As Integer, DayS As Integer, FDayS As String
Dim FDayE As Integer, Mon As Integer, FMon As String
Dim x As Integer
Dim fName As Variant
Dim FilePath As String
Dim smsLastRow As Integer, sms As Worksheet
Dim hoursTracker As Range

On Error GoTo endSub
log.setClass(getClass).setMethod ("NewPayPeriod")

'Ask for password
frmPword.Show

'Deactivate Updating
System.Update False

Set payPeriod = Ranges.getPayPeriodRange()
Set payDay1 = Ranges.getPayDayRange(1)
Set payDay13 = Ranges.getPayDayRange(13)
Set payDay14 = Ranges.getPayDayRange(14)
Set payDay15 = Ranges.getPayDayRange(15)
Set payDay16 = Ranges.getPayDayRange(16)

'Reset Total
log.trace "Clearing Total Misc Range"
System.unprotectSheet Config.getSheet_Total()
Ranges.getTotalMiscRange().ClearContents
Ranges.getTotalMiscRange().ClearComments

For Each ws In ThisWorkbook.Worksheets 'Cycle through each workbook
    'Find the type, which is the text of the first cell
    WsName = ws.name 'Find the abbreviation from the tab
    WsType = PayPeriodTypes.getSheetType(WsName)
    log.trace Words.formatStr("Getting type(\'%s\') of sheet(\'%s\')", WsType, ws.name)
     
    If PayPeriodTypes.isSemimonthlyType(WsType) Then 'Semimonthly type
        Set dataRange = Ranges.getDataRange(WsName) 'Name the Range
        For Each Cell In dataRange
            If Cell.Locked = False Then Cell.ClearContents 'Clear the cell if it's unlocked
        Next Cell
    ElseIf PayPeriodTypes.isMonthlyType(WsType) Or PayPeriodTypes.isInvoiceType(WsType) Then
        System.unprotectSheet ws

        Set dataRange1 = Ranges.getData1Range(WsName) 'define the Range
        Set dataRange2 = Ranges.getData2Range(WsName) 'define the Range
        If Day(payDay1) = 1 Then
            dataRange1.Copy
            dataRange1.PasteSpecial xlPasteValues
            For x = 0 To 12
                ws.Cells(dataRange2.Row + x, dataRange2.Column) = "=PayDay1 + " & x
            Next x
            For x = 13 To 15
                ws.Cells(dataRange2.Row + x, dataRange2.Column) = "=IF(PayDay" & x + 1 & ">0,PayDay" & x + 1 & ","""")"
            Next x
        ElseIf Day(payDay1) = 16 Then
            Set dataRange = Ranges.getDataRange(WsName) 'define the Range
            Set copyFromRange = Ranges.getCopyFromRange(WsName)
            Set pasteToRange = Ranges.getPasteToRange(WsName)
            For Each Cell In dataRange
                If Cell.Locked = False Then Cell.ClearContents 'Clear the cell if it's unlocked
            Next Cell
            For x = 0 To 14
                ws.Cells(dataRange1.Row + x, dataRange1.Column) = "=PayDay1 + " & x
            Next x
            copyFromRange.Copy
            pasteToRange.PasteSpecial Paste:=xlPasteFormulas
            ws.Range(ws.Cells(dataRange2.Row, 1), ws.Cells(dataRange2.Rows(16).Row, 1)).ClearContents
        End If
        System.protectSheet ws
    End If
    
Next ws

'clear OT employee range
'Ranges.getOTEmpRange().ClearContents
    
'Reset Import
System.unprotectSheet Config.getSheet_Import()

payDay1 = WorksheetFunction.Max(Ranges.getPayPeriodRange()) + 1
payDay1.AutoFill Destination:=Ranges.getPayPeriodRange(), Type:=xlFillValues

If Month(payDay13) <> Month(payDay14) Then
    payDay14.ClearContents
End If
If Month(payDay13) <> Month(payDay15) Then
    payDay15.ClearContents
End If
If Month(payDay13) <> Month(payDay16) Then
    payDay16.ClearContents
End If
If Day(payDay16) = 16 Then
    payDay16.ClearContents
End If

Set hoursTracker = Ranges.getHoursSpentTrackerRange()
With Config.getSheet_Import()
    .Range("A:H").ClearContents ' shift planning imported data
    .Range("A:H").Interior.Color = xlNone
    .Range(Words.col(hoursTracker.Columns(1).Column) & ":" & Words.col(hoursTracker.Columns(5).Column)).ClearContents
    .Cells(1, hoursTracker.Columns(1).Column) = "Time Open"
    .Cells(1, hoursTracker.Columns(2).Column) = "Time Saved"
    .Cells(1, hoursTracker.Columns(3).Column) = "Duration"
    .Cells(1, hoursTracker.Columns(4).Column) = "Date"
    .Cells(2, hoursTracker.Columns(4).Column) = Date
    .Cells(2, hoursTracker.Columns(1).Column) = Time
    .Range("H:H").ClearComments

    .Cells(1, 1) = "Employee"
    .Cells(1, 2) = "Date"
    .Cells(1, 3) = "Location"
    .Cells(1, 4) = "Position"
    .Cells(1, 5) = "Start Time"
    .Cells(1, 6) = "End Time"
    .Cells(1, 7) = "Reg Hours"
    .Cells(1, 8) = "Bonus"
End With

'Clear out SMS sheet
Set sms = Config.getSheet_SMS()
If sms.Cells(2, 1).value = vbNullString Then GoTo skipSMS
System.unprotectSheet sms

smsLastRow = sms.Cells(1, 1).End(xlDown).Row
sms.Range("A2:J" & smsLastRow).ClearContents
sms.Range("M:Q").ClearContents
sms.Range("M1") = "Employee"
sms.Range("N1") = "Date"
sms.Range("O1") = "Location"
sms.Range("P1") = "Unmatched Clock Time"
sms.Range("Q1") = "Clock Type"

System.protectSheet sms

skipSMS:
'Figure out File Name
FYear = Year(payDay1) - 2000
DayS = Day(payDay1)
FDayE = Day(WorksheetFunction.Max(payPeriod))
Mon = Month(payDay1)
If DayS < 10 Then
    FDayS = "0" & CStr(DayS)
    Else: FDayS = CStr(DayS)
End If
If Mon < 10 Then
    FMon = "0" & Mon
    Else: FMon = Mon
End If
fName = (FMon & "-" & FDayS & "-" & FYear & " to " & FMon & "-" & FDayE & "-" & FYear)

'Unlock Workbook
For Each ws In ThisWorkbook.Worksheets
    System.unprotectSheet ws
Next

'Re-lock table
Ranges.getImportTableRange().Locked = True

'Go back to Total
Config.getSheet_Total().Activate
Cells.EntireColumn.AutoFit
ActiveWindow.ScrollRow = 1
Cells(1, 1).Select

'Re-lock Workbook
For Each ws In ThisWorkbook.Worksheets
    System.protectSheet ws
Next

System.unprotectSheet Config.getSheet_Import()

'Re-activate Updating
System.Update True

'Save As
FilePath = ThisWorkbook.Path
SkipTracking = True ' This is a global variable meant to stop WorkBook_BeforeSave() from tracking on the import because there's a bug in ThisWorkbook.SaveAs
ThisWorkbook.SaveAs (FilePath & "\" & fName)

'Done! MsgBox
Beep
MsgBox "New Pay Period Saved to " & FilePath, vbOKOnly, "Done!"

System.protectSheet Config.getSheet_Import() 'Hopefully they save after this - wat?

endSub:
If Err <> 0 Then
    log.error Err.Number & " - " & Err.Description
    'MsgBox "Error #: " & Err.Number & " - " & Err.Description, vbCritical
End If

End Sub
