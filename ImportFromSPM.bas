Attribute VB_Name = "ImportFromSPM"
Option Explicit
Option Compare Text

Private Const getClass As String = "ImportFromSPM"

Sub ImportFromSP()

Dim NumTimeCards As Integer
Dim import As Worksheet
Dim NameFix As String
Dim ws As Worksheet
Dim match As Boolean, i As Integer, c As Integer, e As Integer
Dim name As String, H As Variant
Dim EmpRange As Range
Dim location As String
Dim Start As Integer
Dim YorN As Integer
Dim splitName() As String

'Dim PPLoc() As String, PCLoc() As String, PPerr() As String, PCerr() As String, PPe() As Boolean, PCe() As Boolean
'Dim l As Integer, pp As Integer, pc As Integer, UnMatchedmsg As String, msg As Integer

Dim dataRange As Range
Dim Cell As Variant
Dim HasLead As Boolean

Call IsInternetConnected

log.setClass(getClass).setMethod ("ImportFromSP")

Set import = Config.getSheet_Import()

frmPword.Show
FromBeginning:
'Deactivate Updating
System.setToAutomationMode

'if there is import data, erase it all
If import.Cells(2, 1) <> "" Then
    With Range(import.Cells(2, "A").End(xlDown), import.Cells(2, "H"))
        .ClearContents
        .ClearComments
        .Interior.Color = xlNone
    End With
End If

'Import Timecards
Call HumanityAPI.importTimesheets
'Call ShiftLeadFromSP(PPLoc(), PCLoc(), PPerr(), PCerr(), PPe(), PCe(), l, pp, pc, UnMatchedmsg, msg) 'Gets Shift Lead Bonuses based on employees scheduled for same event (assumes exact title equivalency)
Call AddEmployees 'Adds Employees not in 'Total' from Import Sheet after HoursFromSP has run and Deletes from 'Total' the employees who did not work
On Error GoTo endSub

'Find out number of time cards
NumTimeCards = import.Range("A1").End(xlDown).Row - 1

'Sort Fields
import.Sort.SortFields.clear
import.Sort.SortFields.add key:=Range(import.Cells(2, 3), import.Cells(NumTimeCards + 1, 3)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
import.Sort.SortFields.add key:=Range(import.Cells(2, 1), import.Cells(NumTimeCards + 1, 1)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With import.Sort
    .SetRange Range(import.Cells(2, 1), import.Cells(NumTimeCards + 1, 8))
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

For Each ws In ThisWorkbook.Worksheets

    HasLead = False 'Find out if a sheet utilizes a lead by if it has "Lead" somewhere on the sheet
    'HasLead and HasShuttle put the logic in the same place, but I wanted to be pedantic about the naming
    If ws.name <> Config.getSheet_Total().name And ws.name <> Config.getSheet_OT().name And ws.name <> Config.getSheet_Import().name And ws.name <> Config.getSheet_SMS().name Then
        Set dataRange = Range(ws.name & "Data")
        For Each Cell In dataRange
            If Cell.value = ws.name & " Lead" Then
                HasLead = True
                Exit For
            End If
        Next Cell
    End If
    
    If Not HasLead Then
        If (PayPeriodTypes.isMonthlySheet(ws.name) Or PayPeriodTypes.isSemimonthlySheet(ws.name)) Then
            Set EmpRange = Range(ws.name & "Emp")
            location = ws.Cells(1, 1)
            If WorksheetFunction.CountIf(Range(import.Cells(2, 3), import.Cells(NumTimeCards + 1, 3)), location) > 0 Then
                Start = Range(import.Cells(1, 3), import.Cells(1, 3).End(xlDown)).find(What:=location, After:=import.Cells(NumTimeCards + 1, 3), LookIn:=xlFormulas, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False).Row
                For i = Start To NumTimeCards + 1
                    If import.Cells(i, 3) = location Then
                        name = import.Cells(i, 1)
                        match = False
                        For H = EmpRange.Column To EmpRange.Columns(EmpRange.Columns.count).Column
                            If ws.Cells(1, H) = name Then
                                match = True
                                Exit For
                            End If
                        Next H
                        If match = False Then
                            c = EmpRange.Column
                            e = EmpRange.Columns(EmpRange.Columns.count).Column
                            Do Until c > e
                                If ws.Cells(1, c) = "" Then
                                    ws.Cells(1, c) = name
                                    c = e + 1
                                Else
                                    c = c + 1
                                End If
                            Loop
                        End If
                    End If
                Next i
            End If
        End If
    ElseIf HasLead Then
        Set EmpRange = Range(ws.name & "Emp") 'Employee names
        location = ws.Cells(1, 1) 'Account name
        If WorksheetFunction.CountIf(Range(import.Cells(2, 3), import.Cells(NumTimeCards + 1, 3)), location) > 0 Then 'Find out if this location even has time cards
            'Find the first instance of this location
            Start = Range(import.Cells(1, 3), import.Cells(1, 3).End(xlDown)).find(location, import.Cells(NumTimeCards + 1, 3), _
                xlFormulas, xlPart, xlByRows, , False, , False).Row
            For i = Start To NumTimeCards + 1
                If import.Cells(i, 3) = location Then
                    name = import.Cells(i, 1)
                    match = False
                    For H = EmpRange.Column To EmpRange.Columns(EmpRange.Columns.count).Column 'Employee name range start to end
                        If ws.Cells(1, H) = name And ws.Cells(21, H) = import.Cells(i, 4) Then 'This is the part that detects the Lead/Valet
                            match = True
                            Exit For
                        End If
                    Next H
                    If Not match Then
                        c = EmpRange.Column
                        e = EmpRange.Columns(EmpRange.Columns.count).Column
                        Do Until c > e
                            If ws.Cells(1, c) = "" Then
                                ws.Cells(1, c) = name
                                ws.Cells(21, c) = import.Cells(i, 4)
                                c = e + 1
                            Else
                                c = c + 1
                            End If
                        Loop
                    End If
                End If
            Next i
        End If
    End If
Next ws

endSub:

System.restoreToUserMode
Config.getSheet_Import().Activate

If Err <> 0 Then
    MsgBox "Error #: " & Err.Number & " - " & Err.Description, vbCritical
    End
End If
    

End Sub



