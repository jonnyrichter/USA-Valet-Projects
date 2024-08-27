Attribute VB_Name = "HoursFromSPM"
Option Explicit

Sub HoursFromSP() 'Called from ImportFromSP()

Dim bDate As Double, eDate As Double
Dim monthNumber As Long
'Dim pMonth As String (obsolete)
Dim bDay As String, eDay As String
Dim pYear As String
Dim import As Worksheet
Dim r As Integer, numCards As Integer, i As Integer, j As Integer
Dim Extension As String
Dim e As IHTMLElement
Dim resultsTable As IHTMLElement, tableBody As IHTMLElement
Dim headers As IHTMLElementCollection
Dim tableRows As IHTMLElementCollection, tableRow As IHTMLElement
'Dim accountNames() As String, a As Long

Dim empCol As Long, dateCol As Long, locCol As Long, posCol As Long
Dim startCol As Long, endCol As Long, hoursCol As Long

Set import = Config.getSheet_Import()

On Error GoTo endSub

bDate = Ranges.getPayPeriodRange().Rows(1)
eDate = WorksheetFunction.Max(Ranges.getPayPeriodRange())
monthNumber = Month(bDate)
'pMonth = MonthName(monthNumber, True) 'abbreviate month "03/01/17" -> "mar" (obsolete)
bDay = CStr(Day(bDate)) 'cast the day as string (don't really think it's necessary but not worth changing)
pYear = CStr(Year(bDate))
eDay = CStr(Day(eDate))

Extension = "/app/payroll/timesheets/&sdate=" & monthNumber & "/" & bDay & "/" & pYear & _
    "&edate=" & monthNumber & "/" & eDay & "/" & pYear & _
    "&location=-1&t=undefined&ts=undefined&options=undefined&sortby=undefined&remote_site=undefined&openshiftoption=undefined&&min15int=&include_emp_per_pos=-1&include_emp_id=-1&include_emp_eid=-1&wu=undefined&terminal_location=undefined&formatted_times=-1&exclude_disabled_emp=-1&split_overtime_by_rate=undefined&approval_status=undefined&submit=1/"

log.trace Extension
LoginToM.ShiftPlanning Extension

WaitForM.ObjectByClass "ResultsTable", 0
Sleep 3000

Set resultsTable = HTML.getElementsByClassName("ResultsTable")(0)
Set tableBody = resultsTable.getElementsByTagName("tbody")(0)
'Employee
'Date
'Location
'Position
'Start Time
'End Time
'Reg Hours
'Bonus - not brought in Hours Import

Set headers = tableBody.getElementsByTagName("tr")(0) _
     .getElementsByTagName("td")

For i = 0 To headers.length - 1
    If LCase(headers(i).innerText) = "employee" Then
        empCol = i
    End If
    If LCase(headers(i).innerText) = "date" Then
        dateCol = i
    End If
    If LCase(headers(i).innerText) = "location" Then
        locCol = i
    End If
    If LCase(headers(i).innerText) = "position" Then
        posCol = i
    End If
    If LCase(headers(i).innerText) = "start time" Then
        startCol = i
    End If
    If LCase(headers(i).innerText) = "end time" Then
        endCol = i
    End If
    If LCase(headers(i).innerText) = "regular" Then
        hoursCol = i
    End If
Next i

Set tableRows = tableBody.Children
numCards = tableRows.length - 2
'r is row
'i is webPage (pseudo-column) column - It's the column number within the spreadsheet
'j is ACTUAL webPage column

'accountNames = Config.getAllAcountNames()
For r = 2 To numCards + 1
    Set tableRow = tableRows(r - 1)
    
    import.Cells(r, 1) = tableRow.Children(empCol).innerText
    import.Cells(r, 2) = tableRow.Children(dateCol).innerText
    'fix restaurant name 1
    If tableRow.Children(locCol).className = "tti" Then 'elements with "tti" class (no idea) are shortened and we need to get the full name
        If Not tableRow.Children(locCol).hasAttribute("oldtitle") Then
            'the "oldtitle" attribute might not show (fucking Humanity), and if it does, it replaces the "title" attribute (fucking....humanity)
            import.Cells(r, 3) = Strings.Replace(Strings.Replace(tableRow.Children(locCol).getAttribute("title"), "<center>", ""), "</center>", "")
        Else
            import.Cells(r, 3) = tableRow.Children(locCol).getAttribute("oldtitle")
        End If
    Else
        import.Cells(r, 3) = tableRow.Children(locCol).innerText
    End If
    'fix restaurant name 2
    'If Words.endsWith(import.Cells(r, 3), "..") Then 'if has been shortened on humanity
    '    import.Cells(r, 3) = Left(import.Cells(r, 3), Len(import.Cells(r, 3)) - 2) 'remove the '..' from the cell
    '    For a = 0 To UBound(accountNames())
    '        If Words.startsWith(accountNames(a), import.Cells(r, 3)) Then 'if obvioius match
    '            import.Cells(r, 3) = accountNames(a) 'set cell to account name
    '            Exit For
    '        End If
    '    Next a
    'End If
    import.Cells(r, 4) = tableRow.Children(posCol).innerText
    import.Cells(r, 5) = tableRow.Children(startCol).innerText
    import.Cells(r, 6) = tableRow.Children(endCol).innerText
    import.Cells(r, 7) = tableRow.Children(hoursCol).innerText
Next r

endSub:

If Err <> 0 Then
    Selenide.CloseIE
    System.Update True
    MsgBox Err.Number & ": " & Err.Description
    End
End If

End Sub

