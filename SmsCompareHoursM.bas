Attribute VB_Name = "smsCompareHoursM"
Option Explicit
Option Compare Text
Private r As Integer
Private ws As Worksheet

Sub SmsHoursComparison()
'@Eventually make column numbers figured out in script by looping through and counting. Right now their indices are hard-coded.
'Internet Explorer Variables

Dim StartDate As String, endDate As String
Dim msg As String
Dim i As Integer, e As Integer
Dim ImportFirstCell As Range
Dim m As Integer
Dim copied As Boolean

'Time Card Variables
Dim pg As Integer, pgs As Integer, total As Integer
'Dim numEmp As Integer
Dim tableRows As IHTMLElementCollection
Dim tcCountRows As Integer
Dim NumTimeCards As Integer
'Dim tEmp As Collection, tDate As Collection, tLocation As Collection, tLunchIn As Collection, tLunchOut As Collection, tDinnerIn As Collection, tDinnerOut As Collection

Dim tcEmp() As String
Dim tcDate() As String, tcLocation() As String
Dim tcLunchIn() As String, tcLunchOut() As String
Dim tcDinnerIn() As String, tcDinnerOut() As String
Dim valetRange As Range, dateRange As Range, emp As Range
Dim lastRow As Integer
Dim employeeName As String
Dim empFromLastPage As String
Dim In1 As Double, Out1 As Double, In2 As Double, Out2 As Double, In3 As Double, Out3 As Double
Dim twoPM As Double
Dim thirdColumn As String

Dim pagingResult As String
Const pagingResultId As String = "pagingResult"
Const timeClockURL As String = "https://admincorporation1.smsvalet.com/Web/Corporations/Reports.aspx?TYPE=37&C=1"
Const customInput As String = "ctl00_ctl00_CphBodyCommon_CphBodyReport_ReportCtrl_trcTime_CtrlDateTimeType", customValue = "5"
Const startInput As String = "ctl00_ctl00_CphBodyCommon_CphBodyReport_ReportCtrl_trcTime_CtrlDateFrom"
Const endInput As String = "ctl00_ctl00_CphBodyCommon_CphBodyReport_ReportCtrl_trcTime_CtrlDateTo"
Const gridBody As String = "gridBody", gridHead As String = "gridHead"
Dim searchButton(0 To 1) As String: searchButton(0) = "buttonB marginA": searchButton(1) = "0"

On Error GoTo endSub

Set ImportFirstCell = Config.getSheet_Import().Cells(2, 1)

If isEmpty(ImportFirstCell) Or ImportFirstCell.value = vbNullString Then
    MsgBox "Data has not yet been imported from ShiftPlanning", vbCritical, "Too Soon!"
    End
End If

Call IsInternetConnected

frmPword.Show

System.Update False

twoPM = TimeValue("2:00 PM")

copied = copyHoursFromImport()

If Not copied Then
    GoTo endSub
End If

StartDate = CDate(WorksheetFunction.Min(Ranges.getPayPeriodRange()))
endDate = CDate(WorksheetFunction.Max(Ranges.getPayPeriodRange()))

LoginToM.SMSValet '();

IE.navigate timeClockURL 'Go to the report

WaitForM.ObjectById startInput

HTML.getElementById(customInput).value = customValue
HTML.getElementById(startInput).value = StartDate
HTML.getElementById(endInput).value = endDate
HTML.getElementsByClassName(searchButton(0))(CInt(searchButton(1))).Click

WaitForM.ObjectById gridHead
WaitForM.ObjectById pagingResultId

pagingResult = HTML.getElementById(pagingResultId).innerText
pgs = CInt(Split(pagingResult, "/")(1))

Set ws = Config.getSheet_SMS()
If ws.Cells(2, "M") <> vbNullString Or Not isEmpty(ws.Cells(2, "M")) Then
    ws.Range("M2:Q" & ws.Cells(1, "M").End(xlDown).Row).ClearContents
    'ws.Range("I2:J" & ws.Cells(1, "I").End(xlDown).Row).Interior.Color = vbnone
End If
lastRow = ws.Cells(1, 1).End(xlDown).Row
Set valetRange = ws.Range("A2:A" & lastRow)
Set dateRange = Range("B2:B" & lastRow)

For pg = 1 To pgs 'A[1]

    tcCountRows = 0
    total = 0
    
    Set tableRows = HTML.getElementById(gridBody).Children
    tcCountRows = tableRows.length
    
    For Each ele In tableRows 'B[1]
        
        If Words.contains(ele.FirstChild.FirstChild.innerText, "Total") Or Words.contains(ele.Children(10).innerText, "Total") Then
            total = total + 1
        End If

    Next ele 'B[2]
    
    'numEmp = numEmp - grandTotal - totalGreeter - totalSupervisor 'Because "Grand Total", "Total Greeter", and "Total Supervisor" rows
    NumTimeCards = tcCountRows - total 'Subtract the number of rows that say "Total"
    
    ReDim tcEmp(NumTimeCards) As String
    ReDim tcDate(NumTimeCards) As String, tcLocation(NumTimeCards) As String
    ReDim tcLunchIn(NumTimeCards) As String, tcLunchOut(NumTimeCards) As String
    ReDim tcDinnerIn(NumTimeCards) As String, tcDinnerOut(NumTimeCards) As String
    
    e = 0 'THIS IS GOING TO BE A MAINTENANCE NIGHTMARE (NO IT'S NOT - OKAY MAYBE [IT IS]) OH GOD IT FUCKING IS
    For Each ele In tableRows 'D[1]
    
        employeeName = ele.FirstChild.FirstChild.innerText 'This is where Employee's name is grabbed
        thirdColumn = ele.Children(10).innerText
        
        Out3 = SplitTimeFromDate(thirdColumn)
        
        If Not Words.contains(employeeName, "Total") And Not Words.contains(thirdColumn, "Total") Then 'F[1] - Same as Out3 <> -1, just better for readability
            e = e + 1
            '<<<STORE ALL TIMES>>>
            In1 = SplitTimeFromDate(ele.Children(5).innerText) 'Possible for this to be totalGreeter, totalSupervisor, or grandTotal
            Out1 = SplitTimeFromDate(ele.Children(6).innerText)
            In2 = SplitTimeFromDate(ele.Children(7).innerText)
            Out2 = SplitTimeFromDate(ele.Children(8).innerText)
            In3 = SplitTimeFromDate(ele.Children(9).innerText)
            'Out3 = SplitTimeFromDate(thirdColumn) 'Possible for this to be "Total" because reasons - Handled at previous if () { then; }
            
            If employeeName <> vbNullString Then
                tcEmp(e) = Split(employeeName, " ")(1) & " " & Left(employeeName, 2)
            Else
                If e = 1 Then
                    tcEmp(e) = empFromLastPage
                Else
                    tcEmp(e) = tcEmp(e - 1)
                End If
                
            End If
            tcDate(e) = ele.Children(1).innerText
            tcLocation(e) = ele.Children(3).innerText
            
            If In1 < twoPM Then 'H[1]
                tcLunchIn(e) = CTIME(In1)
                tcLunchOut(e) = CTIME(Out1)
                If In2 <> -1 And In2 < twoPM Then 'And In2 > TimeValue(tcLunchIn(e)) Then 'I[1]'The third condition is redundant as any later clock in will be a later time (same for next if)
                    tcLunchIn(e) = vbNullString
                    tcLunchOut(e) = CTIME(Out2)
                    If In3 <> -1 And In3 < twoPM Then 'And In3 > TimeValue(tcLunchIn(e)) Then 'J[1]
                        tcLunchIn(e) = vbNullString
                        tcLunchOut(e) = CTIME(Out3)
                    End If 'J[2]
                ElseIf In2 <> -1 Then 'I[1.5]
                    tcDinnerIn(e) = CTIME(In2)
                    tcDinnerOut(e) = CTIME(Out2)
                    If In3 <> -1 And In3 > TimeValue(tcLunchIn(e)) Then 'K[1]
                        tcDinnerIn(e) = vbNullString
                        tcDinnerOut(e) = CTIME(Out3)
                    End If
                End If 'I[2]
            Else 'H[1.5]
                tcDinnerIn(e) = CTIME(In1)
                tcDinnerOut(e) = CTIME(Out1)
                If In2 <> -1 Then
                    tcDinnerOut(e) = CTIME(Out2)
                    If In3 <> -1 Then
                        tcDinnerOut(e) = CTIME(Out3)
                    End If
                End If
            End If 'H[2]
            
        End If
    Next ele 'D[2]
    i = 0
    For e = 1 To NumTimeCards
        For Each emp In valetRange
            If tcEmp(e) = emp.value And DateValue(emp.Offset(, 1)) = DateValue(tcDate(e)) And Words.contains(emp.Offset(, 2), tcLocation(e), vbTextCompare) Then
                If TimeValue(CDate(emp.Offset(, 4))) < twoPM Then 'Shift is before 2:00 pm and is lunch shift
                    If tcLunchIn(e) <> vbNullString Then
                        emp.Offset(, 6) = roundTime(tcLunchIn(e))
                        tcLunchIn(e) = vbNullString
                    End If
                    If tcLunchOut(e) <> vbNullString Then
                        emp.Offset(, 7) = roundTime(tcLunchOut(e))
                        tcLunchOut(e) = vbNullString
                    End If
                Else 'Shift is after 2:00 pm and is dinner shift
                    If tcDinnerIn(e) <> vbNullString Then
                        emp.Offset(, 6) = roundTime(tcDinnerIn(e))
                        tcDinnerIn(e) = vbNullString
                    End If
                    If tcDinnerOut(e) <> vbNullString Then
                        emp.Offset(, 7) = roundTime(tcDinnerOut(e))
                        tcDinnerOut(e) = vbNullString
                    End If
                End If
            End If
        Next emp
    Next e
    For e = 1 To NumTimeCards
        msg = msg & createMessage(tcEmp(e), tcLunchIn(e), tcDate(e), tcLocation(e), "Lunch In")
        msg = msg & createMessage(tcEmp(e), tcLunchOut(e), tcDate(e), tcLocation(e), "Lunch Out")
        msg = msg & createMessage(tcEmp(e), tcDinnerIn(e), tcDate(e), tcLocation(e), "Dinner In")
        msg = msg & createMessage(tcEmp(e), tcDinnerOut(e), tcDate(e), tcLocation(e), "Dinner Out")
    Next e
    
    empFromLastPage = tcEmp(UBound(tcEmp()))
    If pg < pgs Then
        ClickElementM.ByAttribute "a", "title", "Next" 'Next Arrow (>)
        WaitForM.ObjectToDisappearById "divPreloader" 'waitForLoadingBarToDisappear();
    End If
Next pg 'A[2]
r = 0
finished:

endSub:
On Error Resume Next
InternetHelperM.CloseIE
Call SortSheets
System.Update True

ws.Activate

If Err = 0 Then
    If msg <> vbNullString Then
        m = Words.CharInst(msg, vbNewLine)
        MsgBox "Hours retrieval was a success." & vbNewLine & vbNewLine & "Please check columns M-Q to determine times that don't match and report any possible development mistakes via email to: Richter.Jonathan.R@gmail.com", , "Success! But " & m / 2 & " Shift(s) Not Matched!"
    Else
        MsgBox "Hours retrieval was a success", , "Done!"
    End If
Else
    MsgBox "Error #: " & Err.Number & " - " & Err.Description, vbCritical
    End
End If

End Sub

Private Function copyHoursFromImport() As Boolean

Dim lastRow As Integer, r As Integer
Dim im As Worksheet, sms As Worksheet
Dim spImported As Boolean

On Error GoTo Failure

Set im = Config.getSheet_Import()
Set sms = Config.getSheet_SMS()
System.unprotectSheet sms

If Not isEmpty(sms.Cells(2, 1)) Then
    lastRow = sms.Cells(2, 1).End(xlDown).Row
    sms.Range("A2:J" & lastRow).ClearContents
End If
If Not isEmpty(sms.Cells(2, "M")) Then
    lastRow = sms.Cells(2, "M").End(xlDown).Row
    sms.Range("M2:Q" & lastRow).ClearContents
End If

spImported = im.Cells(2, 1) <> vbNullString
If Not spImported Then
    copyHoursFromImport = False
    Exit Function
End If

lastRow = im.Cells(1, 1).End(xlDown).Row
sms.Range("A2:F" & lastRow).value = im.Range("A2:F" & lastRow).value
For r = 2 To lastRow
    sms.Cells(r, 9) = "=IF(G" & r & "="""",""N/A"",ABS(ROUND((IF(E" & r & "<0.25,E" & r & "+1,E" & r & ")-IF(G" & r & "<0.25,G" & r & "+1,G" & r & "))*24,2)))"
    sms.Cells(r, 10) = "=IF(H" & r & "="""",""N/A"",ABS(ROUND((IF(F" & r & "<0.25,F" & r & "+1,F" & r & ")-IF(H" & r & "<0.25,H" & r & "+1,H" & r & "))*24,2)))"
Next r

'Sort Fields
sms.Sort.SortFields.clear
sms.Sort.SortFields.add key:=Range(sms.Cells(2, 2), sms.Cells(lastRow, 2)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
sms.Sort.SortFields.add key:=Range(sms.Cells(2, 3), sms.Cells(lastRow, 3)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
sms.Sort.SortFields.add key:=Range(sms.Cells(2, 1), sms.Cells(lastRow, 1)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With sms.Sort
    .SetRange Range(sms.Cells(2, 1), sms.Cells(lastRow, 6))
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

sms.Cells.EntireColumn.AutoFit

copyHoursFromImport = True
Exit Function
Failure:
copyHoursFromImport = False

End Function

Public Function roundTime(timeToRound As String) As Double
    roundTime = Round(TimeValue(timeToRound) / (1 / 96)) * (1 / 96)
End Function
Public Function CTIME(timeToConvert As Double) As String
    CTIME = Format(CDate(timeToConvert), "h:mm AM/PM")
End Function
Public Function SplitTimeFromDate(TimeAndDate As String) As Double
If TimeAndDate = vbNullString Or Words.contains(TimeAndDate, "Total") Then
    SplitTimeFromDate = -1
Else
    SplitTimeFromDate = TimeValue(Split(TimeAndDate, vbNewLine)(0))
End If
End Function
Private Function createMessage(employee As String, clockTime As String, clockDate As String, location As String, clockType As String)
If clockTime <> vbNullString Then
    r = r + 1
    createMessage = vbNewLine & employee & " hours not found for " & roundTime(clockTime) & " on " & clockDate & " at " & location
    ws.Cells(r + 1, "M") = employee
    ws.Cells(r + 1, "N") = clockDate
    ws.Cells(r + 1, "O") = location
    ws.Cells(r + 1, "P") = roundTime(clockTime)
    ws.Cells(r + 1, "Q") = clockType
Else
    createMessage = vbNullString
End If
End Function

