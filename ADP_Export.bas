Attribute VB_Name = "ADP_Export"
Option Explicit
Option Compare Text
'DISABLE FULL SCREEN DURING TESTING

'Find an alternative to SendKeys - Yet to Find

Sub ExportToADP()

Dim yn As Boolean

'Make sure to enable "Microsoft Internet Controls", & "Microsoft HTML Object Library" in 'References
Dim HTMLname As String, tokens() As String
Dim ADPname As String
'Payroll variables
Dim total As Worksheet
Dim employee As Variant

Dim RegHours As String, ccTips As String, Reimb As String, otHours As String
Dim pgs As Integer, pg As Integer, numEmp As Integer
Dim emp() As String, match() As Boolean
Dim e As Integer, i As Integer
Dim strMissingEmp As String
'Checker variables
Dim pgs_() As String 'For pulling the last page number from "1 of 4"
Dim payHours As String, payTips As String, payReimb As String, payOT As String
Dim adphours As String, adptips As String, adpreimb As String, adpOT As String
Dim difHours As String, difTips As String, difReimb As String, difOT As String
Dim headerText As String

Dim regHoursCol As Integer, ccTipsCol As Integer, reimbCol As Integer, otHoursCol As Integer
Dim regHours_input As IHTMLElement, ccTips_input As IHTMLElement, mileageReimb_input As IHTMLElement, overtimeHours_input As IHTMLElement

On Error GoTo endSub

log.setClass("ADP_Export").setMethod ("ExportToADP")

If VersionControl.TestStatus Then
    yn = MsgBox("Payroll is in test status, it will not save. Continue?", vbYesNo, "Test Status") = vbYes
    If Not yn Then End
End If

For Each employee In Ranges.getTotalEmpRange()
    If employee <> "" And employee.Offset(0, 1) = "" Then MsgBox "Missing Regular Wage for " & employee, vbCritical, "Wage Needed For Proper Gratuity Calculation": Exit Sub
    If employee <> "" And employee.Offset(0, 2) = "" Then MsgBox "Missing Secondary Wage for " & employee, vbCritical, "Wage Needed For Proper Gratuity Calculation": Exit Sub
Next employee

If Ranges.getTotalVarianceRange().value <> 0 Then
    If MsgBox("There is a variance of " & Ranges.getTotalVarianceRange() & ". Are you absolutely sure you'd like to proceed?", vbYesNo, "Warning!") = vbNo Then Exit Sub
End If

Call IsInternetConnected

frmPword.Show

MsgBox "Please refrain from touching the keyboard or mouse during export", vbInformation, "Important!"

Set total = Config.getSheet_Total()

numEmp = Ranges.getTotalEmpRange().Rows.count
ReDim match(numEmp) As Boolean
ReDim emp(numEmp) As String
For e = 1 To numEmp
    emp(e) = Ranges.getTotalEmpRange().Rows(e)
Next e

LoginToM.ADP  'Logs in to ADP to main menu screen'@TheatreMode True/False
If VersionControl.TestStatus Then
    IE.FullScreen = True
End If

HTML.getElementById("PAYRUN_REGULAR").Click

WaitForM.ObjectByVisibleText "option", "Semimonthly"

HTML.getElementsByClassName("idCboSelectFrequency")(0).value = "S" 'Select Semimonthly
HTML.getElementsByClassName("enterpayroll")(1).Click

WaitForM.ObjectByClass "dgrid-scroller"

pgs_() = Split((HTML.getElementsByClassName("dgrid-status")(0).innerText), " of ")
pgs = CInt(Replace(pgs_(1), ":", ""))

i = 0
Const colsFromFirst As Long = 5
For Each ele In HTML.getElementsByClassName("dgrid-cell-container")
    i = i + 1
    'minus 5 because...?
    headerText = ele.innerText
    If Words.contains(headerText, "Regular") And Words.contains(headerText, "Hours") Then
        regHoursCol = i - colsFromFirst
    End If
    If Words.contains(headerText, "CC") And Words.contains(headerText, "Tips") And Words.contains(headerText, "Owed") Then
        ccTipsCol = i - colsFromFirst
    End If
    If Words.contains(headerText, "Mileage") And Words.contains(headerText, "Reimb") Then
        reimbCol = i - colsFromFirst
    End If
    If Words.contains(headerText, "Overtime") And Words.contains(headerText, "Hours") Then
        otHoursCol = i - colsFromFirst
    End If
    'end loop if all headers found
    If regHoursCol > 0 And ccTipsCol > 0 And reimbCol > 0 And otHoursCol > 0 Then
        Exit For
    End If
Next ele
i = 0

For pg = 1 To pgs
    'Go to the next page (1st page will create no action)
    If pg > 1 Then
        For Each ele In HTML.getElementsByClassName("dgrid-page-link")
            If CInt(ele.innerText) = pg Then
                ele.Click
                Exit For
            End If
        Next ele
        WaitForM.ObjectToHaveText HTML.getElementsByClassName("dgrid-status")(0), "Page " & pg & " of " & pgs & ":"
    End If
    For Each ele In HTML.getElementsByClassName("dgrid-cell dgrid-cell-padding dgrid-column-1-0-" & regHoursCol & " field-EARN_REG_EeEarnQty fieldCode_EeEarnQty dgrid-input-editable dgrid-cell-align-right")
        e = 0
        HTMLname = Split(ele.getAttribute("data-grid-row-title"), "|")(0)
        If Words.CharInst(HTMLname, " ") = 2 Then
            ADPname = Split(HTMLname, " ")(0) & " " & Split(HTMLname, " ")(2)
        Else
            ADPname = HTMLname
        End If
        'Convert 'first last' to 'last, first'
        tokens() = Strings.Split(ADPname, " ")
        ADPname = tokens(UBound(tokens())) & ", " & tokens(0)
        For Each employee In Ranges.getTotalEmpRange()
            e = e + 1
            If employee = "" Then
                GoTo NextEle
            End If
            
            RegHours = total.Cells(employee.Row, "D")
            ccTips = total.Cells(employee.Row, "E")
            Reimb = total.Cells(employee.Row, "F")
            otHours = total.Cells(employee.Row, "G")
            
            'works with the first match but not after that.
            If ADPname = employee Then
                match(e) = True
                'As of 5/18/(year?) there was a problem here. CC Tips were getting entered into the salary column
                'I suspect this has to do with sending tab to the browser because it seems the click didn't have time to hit the next cell before sending the value to the cell
                'A possible solution would be to use a workaround so that the correct cell has time to be clicked before the value is set
                'The IDEAL solution would be to stop using sendKeys - <- This fuckin' guy LOL
                If RegHours <> 0 Then
                    Call enterValue(ele.parentElement.Children(regHoursCol), RegHours)
                End If
                If ccTips <> 0 Then
                    Call enterValue(ele.parentElement.Children(ccTipsCol), ccTips)
                End If
                If Reimb <> 0 Then
                    Call enterValue(ele.parentElement.Children(reimbCol), Reimb)
                End If
                If otHours <> 0 Then
                    Call enterValue(ele.parentElement.Children(otHoursCol), otHours)
                End If
                
                Exit For
            End If
        Next employee
NextEle:
    Next ele
Next pg

'Match ADP totals to Payroll totals
payHours = Format(Ranges.getTotalTotalsRange().Columns(1), "Standard")
payTips = Format(Ranges.getTotalTotalsRange().Columns(2), "Currency")
payReimb = Format(Ranges.getTotalTotalsRange().Columns(3), "Currency")
payOT = Format(Ranges.getTotalTotalsRange().Columns(4), "Currency")

'from adp
Dim row_grid As IHTMLElement
Set row_grid = HTML.getElementsByClassName("dgrid-column-totals-row dgrid-row")(0)
adphours = Format(Mid(row_grid.getElementsByClassName("dgrid-cell dgrid-cell-padding dgrid-column-1-0-" & regHoursCol & " field-EARN_REG_EeEarnQty fieldCode_EeEarnQty dgrid-cell-align-right")(0).innerText, 2), "Standard")
adptips = Format(Mid(row_grid.getElementsByClassName("dgrid-cell dgrid-cell-padding dgrid-column-1-0-" & ccTipsCol & " field-EARN_CREDTIPP_EeEarnAmt fieldCode_EeEarnAmt dgrid-cell-align-right")(0).innerText, 2), "Currency")
adpreimb = Format(Mid(row_grid.getElementsByClassName("dgrid-cell dgrid-cell-padding dgrid-column-1-0-" & reimbCol & " field-EARN_MILREINT_EeEarnAmt fieldCode_EeEarnAmt dgrid-cell-align-right")(0).innerText, 2), "Currency")
adpOT = Format(Mid(row_grid.getElementsByClassName("dgrid-cell dgrid-cell-padding dgrid-column-1-0-" & otHoursCol & " field-EARN_OVT_EeEarnQty fieldCode_EeEarnQty dgrid-cell-align-right")(0).innerText, 2), "Currency")

'difference
difHours = Format(CDbl(adphours) - CDbl(payHours), "Standard")
difTips = Format(CDbl(adptips) - CDbl(payTips), "Currency")
difReimb = Format(CDbl(adpreimb) - CDbl(payReimb), "Currency")
difOT = Format(CDbl(adpOT) - CDbl(payOT), "Currency")

'Save ADP
If VersionControl.TestStatus Then
    Stop
Else
    For Each ele In HTML.getElementsByTagName("a")
        If ele.Title = "Save" And ele.getAttribute("rel") = "grey" And ele.innerText = "Save" Then
            ele.Click
        End If
    Next ele
End If

WaitForM.BrowserToLoad

'Take down full screen
IE.TheaterMode = False
IE.Top = 0
IE.Left = 0
IE.Height = 1000
IE.Width = 1000

'Display missing employees from ADP's database
i = 0
For e = 1 To numEmp
    If match(e) = False And emp(e) <> "" Then
        i = i + 1
        emp(e) = Mid(emp(e), InStr(1, emp(e), " ") + 1) & " " & Left(emp(e), InStr(1, emp(e), ",") - 1)
        If i = 1 Then
            strMissingEmp = emp(e)
        ElseIf i > 1 Then
            strMissingEmp = strMissingEmp & vbNewLine & emp(e)
        End If
    End If
Next e


If i >= 1 Then
    MsgBox "Employees not in ADP's system:" & vbNewLine & vbNewLine & strMissingEmp, vbInformation, "Missing Employees"
End If

MsgBox "ADP Hours: " & adphours & vbNewLine & "Payroll Hours: " & payHours & vbNewLine & "----------" & vbNewLine & "Difference: " & difHours & vbNewLine & vbNewLine _
    & "ADP Tips: " & adptips & vbNewLine & "Payroll Tips: " & payTips & vbNewLine & "----------" & vbNewLine & "Difference: " & difTips & vbNewLine & vbNewLine _
    & "ADP Reimb.: " & adpreimb & vbNewLine & "Payroll Reimb.: " & payReimb & vbNewLine & "----------" & vbNewLine & "Difference: " & difReimb & vbNewLine & vbNewLine _
    & "ADP OT: " & adpOT & vbNewLine & "Payroll OT: " & payOT & vbNewLine & "----------" & vbNewLine & "Difference: " & difOT & vbNewLine & vbNewLine _
    , vbInformation, "Comparison Check!"

MsgBox "Exportation to ADP complete", , "Done!"

endSub:
If Err <> 0 Then
    MsgBox "There was an error." & vbNewLine & Err.Number & ": " & Err.Description, vbCritical, "ERROR!"
End If
InternetHelperM.CloseIE

End Sub

Private Sub enterValue(ele2 As IHTMLElement, val As String)
    ele2.Click
    HTML.activeElement.value = val
    UserKeys.sendTab
    Sleep 500
End Sub
