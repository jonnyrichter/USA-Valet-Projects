VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const getClass As String = "ThisWorkbook"

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Dim yn As Integer
If TestStatus Then
    yn = MsgBox("Leave Workbook in Test Mode?", vbYesNo, "Test Mode On!")
    If yn = vbNo Then
        Cancel = True
    End If
End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean) 'Make this more dynamic at the formula

'Debug.Print Strings.Split(ThisWorkbook.Path(), "\")(2) @TODO
'Jonathan

Dim import As Worksheet
Dim hoursTracker As Range
Dim user As String
Dim endTimeRange As Range

On Error Resume Next
user = Environ("username")

log.setClass(getClass).setMethod ("Workbook_BeforeSave")

If SkipTracking Then GoTo SkipThis

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Set hoursTracker = Ranges.getHoursSpentTrackerRange()

Set import = Config.getSheet_Import()

System.unprotectSheet import

If import.Cells(2, hoursTracker.Columns(2).Column) = Empty Then 'fill first time/user slot
    import.Cells(2, hoursTracker.Columns(2).Column) = Time
    import.Cells(2, hoursTracker.Columns(5).Column) = user
Else 'fill last time/user slot
    Set endTimeRange = import.Cells(1, hoursTracker.Columns(2).Column).End(xlDown).Offset(1, 0)
    endTimeRange.value = Time
    endTimeRange.Offset(0, 3).value = user
End If

If import.Cells(2, hoursTracker.Columns(3).Column) = Empty Then 'fill first formula slot
    import.Cells(2, hoursTracker.Columns(3).Column) = "=HRS(" & Words.col(hoursTracker.Columns(1).Column) & "2," & Words.col(hoursTracker.Columns(2).Column) & "2)"
Else 'fill last formula slot
    import.Cells(1, hoursTracker.Columns(3).Column).End(xlDown).Offset(1, 0) = _
        "=HRS(" & Words.col(hoursTracker.Columns(1).Column) & import.Cells(1, Words.col(hoursTracker.Columns(2).Column)).End(xlDown).Row & "," & _
        Words.col(hoursTracker.Columns(2).Column) & import.Cells(1, Words.col(hoursTracker.Columns(1).Column)).End(xlDown).Row & ")"
End If '^Needs to be more dynamic

If Not isEmpty(import.Cells(2, hoursTracker.Column)) Then
    import.Cells(1, hoursTracker.Columns(4).Column).End(xlDown).Offset(1, 0) = Date
    import.Cells(1, hoursTracker.Columns(1).Column).End(xlDown).Offset(1, 0) = Time
Else
    import.Cells(2, hoursTracker.Columns(4).Column) = Date
    import.Cells(2, hoursTracker.Columns(1).Column) = Time
End If

import.Cells.EntireColumn.AutoFit

System.protectSheet import

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

SkipThis:
SkipTracking = False
If Err.Number <> 0 Then log.error Err.Number & ": " & Err.Description

End Sub

Private Sub Workbook_Open() 'make this more dynamic

Dim import As Worksheet
Dim hoursTracker As Range

On Error Resume Next
log.setClass(getClass).setMethod ("Workbook_Open")

If TestStatus Then
    MsgBox "Workbook is currently in test mode.", vbInformation, "Test Mode On!"
End If

System.Update False

Set import = Config.getSheet_Import()

Set hoursTracker = Ranges.getHoursSpentTrackerRange()

System.unprotectSheet import

If import.Cells(1, hoursTracker.Columns(1).Column).End(xlDown).Row <> import.Cells(1, Words.col(hoursTracker.Columns(2).Column)).End(xlDown).Row Then
    import.Cells(1, hoursTracker.Columns(1).Column).End(xlDown).ClearContents
    import.Cells(1, hoursTracker.Columns(4).Column).End(xlDown).ClearContents
End If

If Not isEmpty(import.Cells(2, hoursTracker.Columns(4).Column)) Then
    import.Cells(1, hoursTracker.Columns(4).Column).End(xlDown).Offset(1, 0) = Date
    import.Cells(1, hoursTracker.Columns(1).Column).End(xlDown).Offset(1, 0) = Time
Else
    import.Cells(2, hoursTracker.Columns(4).Column) = Date
    import.Cells(2, hoursTracker.Columns(1).Column) = Time
End If

import.Cells.EntireColumn.AutoFit

System.protectSheet Config.getSheet_Import()

System.Update True

If Err.Number <> 0 Then log.error Err.Number & ": " & Err.Description

End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

Dim ws As Worksheet
Dim WSEmp As Range, WSSortField As Range
Dim fRow As Integer, lRow As Integer
Dim sCol As Integer

On Error GoTo endSub
Set ws = ThisWorkbook.ActiveSheet
If PayPeriodTypes.isMonthlySheet(ws.name) Or PayPeriodTypes.isSemimonthlySheet(ws.name) Then
'If ActiveSheet.Cells(1, 1).Comment.text = "M" Or ActiveSheet.Cells(1, 1).Comment.text = "S" Then

    Set WSSortField = Range(ws.name & "SortField")
    Set WSEmp = Range(ws.name & "Emp")
    
    If Selection.Column = WSEmp(WSEmp.Row, WSEmp.Columns.count).Column And Selection.Row = WSEmp.Row And Selection.Offset(0, -1) <> "" Then
        sCol = Selection.Column: fRow = WSEmp.Row: lRow = WSSortField(WSSortField.Rows.count, Selection.Column).Row
        System.unprotectSheet ws
        Range(ws.Cells(fRow, sCol), ws.Cells(lRow, sCol)).Insert xlToRight
        Range(ws.Cells(fRow, sCol + 1), ws.Cells(lRow, sCol + 1)).Copy
        Range(ws.Cells(fRow, sCol + 1), ws.Cells(lRow, sCol)).PasteSpecial xlPasteFormulas
        ws.Cells.EntireColumn.AutoFit
        System.protectSheet ws
    End If
End If

endSub:
System.protectSheet ws

End Sub
