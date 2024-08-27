Attribute VB_Name = "ReimbursementM"
Option Explicit
Sub Reimbursement()
Attribute Reimbursement.VB_ProcData.VB_Invoke_Func = "r\n14"

Dim numReimb As Integer, r As Integer
Dim location() As String
Dim TravelTime() As Variant
Dim employee As String, Cell As Variant
Dim TH As Worksheet
Dim dy As Integer, mnth As Integer, yr As Integer
Dim found As Boolean, Extension As String

On Error GoTo endSub

'But... Why?
MsgBox "This functionality has been discontinued until further notice", vbCritical, "Disabled!"
End

If ThisWorkbook.ActiveSheet.name <> "PP" And ActiveSheet.name <> "PC" Then MsgBox "Try Again", vbCritical, "Wrong Cell": Exit Sub
If ThisWorkbook.ActiveSheet.Cells(Selection.Row, 2) <> "Reimb. Travel Minutes" Then MsgBox "Try Again", vbCritical, "Wrong Cell": Exit Sub
If ThisWorkbook.ActiveSheet.Cells(1, Selection.Column) = "" Then MsgBox "Try Again", vbCritical, "Wrong Cell": Exit Sub

Call IsInternetConnected

Application.Cursor = xlWait

Set TH = Config.getSheet_Total()

employee = Cells(1, Selection.Column)
For Each Cell In Ranges.getEmployeeListRange()
    If Cell.value = employee Then employee = TH.Cells(Cell.Row, "A")
Next Cell

dy = Day(Cells(Selection.Row - 3, "A"))
mnth = Month(Cells(Selection.Row - 3, "A")) - 1 'minus one because Shift Planning does it that way for some reason
yr = Year(Cells(Selection.Row - 3, "A"))

Extension = "/app/schedule/list/day/schedule/868658%2c500692%2c340675%2c340674%2c340673%2c868659%2c500691%2c340672/" & dy & "%2c" & mnth & "%2c" & yr

LoginToM.ShiftPlanning Extension

WaitForM.ObjectById "shiftlist"

For Each ele In HTML.getElementsByClassName("fourth")
    'log.trace @todo, don't need this ELE.FirstChild.innerText
    If Words.contains(ele.FirstChild.innerText, "Private Party") Or Words.contains(ele.FirstChild.innerText, "Parking Control") Then
        If Words.contains(ele.PreviousSibling.innerText, employee) Then
            numReimb = numReimb + 1
        End If
    End If
Next ele

If numReimb = 0 Then MsgBox "Could not find shift for selected employee", , "No shifts found": IE.Quit: GoTo endSub

ReDim location(numReimb) As String
ReDim TravelTime(numReimb) As Variant

For Each ele In HTML.getElementsByClassName("fourth")
    If Words.contains(ele.FirstChild.innerText, "Private Party") Or Words.contains(ele.FirstChild.innerText, "Parking Control") Then
        If Words.contains(ele.PreviousSibling.innerText, employee) Then
            r = r + 1
            location(r) = ele.Children(2).innerText
        End If
    End If
Next ele

For r = 1 To numReimb
IE.Quit
    location(r) = Replace(location(r), " ", "+")
    url = "https://www.google.com/maps/dir/800+J+St,+Sacramento,+CA+95814/" & location(r)
    
    InternetHelperM.OpenIE url
    
    WaitForM.ObjectByVisibleText "button", "DETAILS"
    
    For Each ele In HTML.getElementsByTagName("span")
        If Words.contains(ele.innerText, "without traffic") Then
            TravelTime(r) = ele.Children(0).innerText
            Exit For
        End If
    Next ele
    
    TravelTime(r) = Replace(TravelTime(r), "h", "|")
    TravelTime(r) = Trim(Replace(TravelTime(r), "min", ""))
    TravelTime(r) = CInt(Split(TravelTime(r), "|")(0)) * 60 + CInt(Split(TravelTime(r), "|")(1))
    
    If TravelTime(r) >= 30 And TravelTime(r) >= Selection Then
        Selection = TravelTime(r)
        found = True
    End If
Next r

IE.Quit
Beep
If found = False Then MsgBox "Travel time is either less than 30 minutes or currently designated reimbursement time.", , "Did Not Reimburse"

endSub:
Application.Cursor = xlDefault

If Err <> 0 Then
    Beep
    MsgBox "Error #: " & Err.Number & " - " & Err.Description, vbCritical
    System.Update True
    End
End If

End Sub

