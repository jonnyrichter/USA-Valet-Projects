Attribute VB_Name = "pullCommentsFromShiftPlanningM"
Option Explicit
Option Compare Text

Sub PullCommentsFromShiftPlanning() 'This will only work after comment has been hovered over. Tricky.

Dim ldy As String, fdy As String, mnth As String, yr As Integer
Dim mnthName As String
Dim i As Integer
Dim com As Worksheet, Extension As String
Set com = ThisWorkbook.Worksheets("Comments")

'On Error GoTo endSub

Call IsInternetConnected

'frmPword.Show

If Day(WorksheetFunction.Min([payPeriod])) < 10 Then
    fdy = "0" & Day(WorksheetFunction.Min([payPeriod]))
Else
    fdy = "" & Day(WorksheetFunction.Min([payPeriod]))
End If
ldy = Day(WorksheetFunction.Max([payPeriod])) & ""

mnth = Month(WorksheetFunction.Min(Ranges.getPayPeriodRange()))  'Minus one because Shift Planning does it that way for some reason
mnthName = MonthName(mnth, True)
yr = Year(WorksheetFunction.Min(Ranges.getPayPeriodRange()))

'Extension = "/app/timeclock/manage/sdate%253D" & mnthName & "%2520" & fdy & "%252C%2520" & yr & "%2526edate%253D" & mnthName & "%2520" & ldy & "%252C%2520" & yr & "%2526status%253Dapproved%2526s%253D-1%2526e%253D-1%2526l%253Dundefined%2526combine%253Dundefined%2526skill%253D-1/"
'
'LoginToM.ShiftPlanning Extension
'
'i = 0
'For Each ELE In HTML.getElementsByClassName("qtip-content qtip-content")
'    i = i + 1
'    COM.Cells(i, "A") = ELE.innerText
'Next ELE
'
'endSub:
'If Err <> 0 Then
'    MsgBox "Error#: " & Err & Err.Description, vbCritical
'    Exit Sub
'End If

End Sub
