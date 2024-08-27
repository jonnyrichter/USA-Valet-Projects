Attribute VB_Name = "ShiftLeadFromSPM"
Option Explicit
Option Compare Text
'Option Base 0 'I should really do this but I don't care enough

Sub ShiftLeadFromSP(PPLoc() As String, PCLoc() As String, PPerr() As String, PCerr() As String, PPe() As Boolean, PCe() As Boolean, l As Integer, _
    pp As Integer, pc As Integer, UnMatchedmsg As String, msg As Integer)   'Called from HoursFromSP()

Dim ele2 As IHTMLElement

Dim fDay As Integer, lDay As Integer, yr As Integer, mnth As Integer, mnthName As String
Dim currentMonth As Long, currentDay As Long

Dim NumPPLeads As Integer, NumPCLeads As Integer
Dim PPLead() As String, PCLead() As String
Dim NumValet() As Integer, NumAttendant() As Integer
Dim PPDate() As String, PCDate() As String, tempLead() As String
Dim employee As Variant, strLead As String
Dim p As Integer, c As Integer, i As Integer, ClassDay As Integer
Dim dateClass As String
Dim Errored As Boolean, Extension As String

Dim eventLocation As String, EventDate As String
Dim leadLocation As String
Dim EventType As String, LeadType As String
Dim match As Boolean, PossibleMatch As String, LeadLoc As Variant
Dim employeeName As String
Dim Managers() As String, m As Integer, ManagerReplace As Variant, WithoutManager As String
Dim im As Worksheet, Bonus As Integer
Dim ManagerRange As Range

Const ppLeadName = "PP Lead", pcLeadName = "PC Lead"

On Error GoTo endSub

currentMonth = CLng(Format(Date, "m")) 'get today's date to compare to last day of the month
currentDay = CLng(Format(Date, "d"))

Set ManagerRange = Ranges.getManagerRange()
ReDim Managers(ManagerRange.Rows.count - 1)
For m = 0 To ManagerRange.Rows.count - 1
    Managers(m) = ManagerRange.Cells(m + 1, ManagerRange.Column)
Next m

lDay = Day(WorksheetFunction.Max(Ranges.getPayPeriodRange())) 'last day
fDay = Day(WorksheetFunction.Min(Ranges.getPayPeriodRange())) 'first day

mnth = Month(WorksheetFunction.Min(Ranges.getPayPeriodRange())) - 1 'minus one because Shift Planning does it that way for some reason

yr = Year(WorksheetFunction.Min(Ranges.getPayPeriodRange()))
'%2c = ",", the numbers themselves are the schedule IDs of the PP/PC shifts (lead, valet, and attendant)
Extension = "/app/schedule/list/month/schedule/340675%2c868658%2c500692%2c340674%2c868659%2c500691%2c340673%2c340672/" & lDay & "%2c" & mnth & "%2c" & yr

log.trace Extension

mnth = mnth + 1 'reset back to the way normal people do it

'Don't go past today's date because you the shift lead might not be deteremined for future dates
'This stays here because mnth = mnth + 1 resets Apr to 4 like it should be
If currentMonth = mnth And currentDay < lDay Then
    lDay = currentDay
End If

mnthName = MonthName(mnth, True)

LoginToM.ShiftPlanning Extension
Sleep 5000
WaitForM.ObjectById "shiftlist"

'Count the number of PP and PC Leads
For ClassDay = fDay To lDay
    
    dateClass = getClassForPayPeriod(mnth, ClassDay, yr)
    For Each ele In HTML.getElementsByClassName(dateClass)
        If ele.Children.length > 1 Then 'is not header
            employeeName = ele.Children(2).innerText
            WithoutManager = employeeName
            For Each ManagerReplace In Managers()
                WithoutManager = Replace(WithoutManager, ManagerReplace, vNil)
            Next ManagerReplace
            WithoutManager = Replace(WithoutManager, vbNewLine, vNil)
            
            If Not VBA.isEmpty(WithoutManager) Then 'There was more than managers working
                strLead = employeeName
                For m = 0 To UBound(Managers())
                    strLead = Replace(strLead, Managers(m), vNil)
                Next m
                Select Case getLeadType(ele)
                    Case ppLeadName '="PP Lead"
                        NumPPLeads = NumPPLeads + countEmployees(strLead)
                        GoTo NextEvent
                    Case pcLeadName '="PC Lead"
                        NumPCLeads = NumPCLeads + countEmployees(strLead)
                End Select
            End If
        End If
NextEvent:
    Next ele
Next ClassDay

ReDim PPLead(NumPPLeads) As String, PCLead(NumPCLeads) As String
ReDim NumValet(NumPPLeads) As Integer, NumAttendant(NumPCLeads) As Integer
ReDim PPDate(NumPPLeads) As String, PCDate(NumPCLeads) As String
ReDim PPerr(NumPPLeads) As String, PCerr(NumPCLeads) As String
ReDim PPe(NumPPLeads) As Boolean, PCe(NumPCLeads) As Boolean
ReDim PPLoc(NumPPLeads) As String, PCLoc(NumPCLeads) As String

'Actually count the employees that lead led
p = 1: c = 1: pp = 0: pc = 0: l = 0: UnMatchedmsg = "": msg = 0
For ClassDay = fDay To lDay 'First day of pay period to last
    
    dateClass = getClassForPayPeriod(mnth, ClassDay, yr)
    For Each ele In HTML.getElementsByClassName(dateClass) 'Collection of the pay period, one date at a time
        If ele.Children.length > 1 Then 'There is actually a shift
            employeeName = ele.Children(2).innerText 'className='third'
            WithoutManager = employeeName
            For Each ManagerReplace In Managers()
                WithoutManager = Replace(WithoutManager, ManagerReplace, vbNullString) 'taken care of, replaces managers like steve and brian from employees
            Next ManagerReplace
            WithoutManager = Replace(WithoutManager, vbNewLine, vbNullString) 'Get rid of extra \n
            
            If Not isEmpty(WithoutManager) Then

                Select Case getLeadType(ele) 'Check if PP Lead or PC Lead
                    Case ppLeadName 'Is Private Party Lead
                        Errored = False
                        pp = pp + p
                        PPLead(pp) = ele.Children(2).innerText
                        For m = 0 To UBound(Managers())
                            PPLead(pp) = Replace(PPLead(pp), Managers(m), vbNullString)
                            'PPLead(pp) = Replace(PPLead(pp), vbNewLine & Managers(m), vbNullString)
                        Next m
                        PPDate(pp) = mnthName & " " & ClassDay
                        PPLoc(pp) = Split(ele.Children(3).Children(0).innerText, vbNewLine)(0)
                        p = countEmployees(PPLead(pp))
                        
                        If p > 1 Then 'if there is more than one lead
                            Errored = True
                            tempLead() = Split(PPLead(pp), vbNewLine)
                            For i = 0 To p - 1
                                PPLead(pp + i) = tempLead(i)
                                PPe(pp + i) = True
                                PPDate(pp + i) = mnthName & " " & ClassDay
                                PPLoc(pp + i) = ele.Children(3).Children(2).innerText
                                For Each ele2 In HTML.getElementsByClassName(dateClass)
                                    If ele2.Children.length > 1 Then
                                        'has valets
                                    If Not ele2.Children(2).innerText = vbNullString Then 'Has valets?
                                        'the title matches
                                        If getLocation(ele2) = getLocation(ele) Then 'NumValet = NumValet + {} because multiple PP shifts for one Lead
                                                If getLeadType(ele2) <> ppLeadName Then
                                                    NumValet(pp + i) = NumValet(pp + i) + countEmployees(ele2.Children(2).innerText)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next ele2
                                PPerr(pp) = PPerr(pp) & vbNewLine & tempLead(i)
                            Next i
                            PPerr(pp) = "Specified more than 1 Non-Steven/Brian Lead for Private Party." & vbNewLine & vbNewLine & "Location:" & vbNewLine & PPLoc(pp) & vbNewLine & vbNewLine & "Date:" & vbNewLine & PPDate(pp) & vbNewLine & vbNewLine _
                                & "Shift Leads:" & PPerr(pp) & vbNewLine & vbNewLine & "Leads granted $15 bonus highlighted in yellow. If incorrect result, please designate ONE Lead on Shift Planning and try again OR edit bonuses manually."
                        Else 'p is not greater than 1
                            For Each ele2 In HTML.getElementsByClassName(dateClass)
                                If ele2.Children.length > 1 Then 'Ignore the header, which will have a children length of 1
                                    'has valets
                                    If Not ele2.Children(2).innerText = vbNullString Then
                                        'the title matches
                                        If getLocation(ele2) = getLocation(ele) Then 'NumValet = NumValet + {} because multiple PP shifts for one Lead
                                            If getLeadType(ele2) <> ppLeadName Then
                                                NumValet(pp) = NumValet(pp) + countEmployees(ele2.Children(2).innerText)
                                            End If
                                        End If
                                    End If
                                End If
                            Next ele2
                        End If
                        If NumValet(pp) <= 0 Then
                            PPe(pp) = True
                            If Errored = False Then
                                PPerr(pp) = "No Non-Lead valets found for Private Party." & vbNewLine & vbNewLine & "Location:" & vbNewLine & PPLoc(pp) & vbNewLine & vbNewLine & "Date:" & vbNewLine & PPDate(pp) & vbNewLine & vbNewLine & "Shift Lead:" _
                                    & vbNewLine & PPLead(pp) & vbNewLine & vbNewLine & "Lead granted $15 bonus highlighted in yellow. If incorrect result, please make correction and try again OR edit bonus manually."
                            Else
                                PPerr(pp) = "No Non-Lead valets found for Private Party." & vbNewLine & vbNewLine & "Location:" & vbNewLine & PPLoc(pp) & vbNewLine & vbNewLine & "Date:" & vbNewLine & PPDate(pp) & vbNewLine & vbNewLine & "Shift Leads:" _
                                    & vbNewLine & PPLead(pp) & vbNewLine & vbNewLine & "Lead granted $15 bonus highlighted in yellow. If incorrect result, please make correction and try again OR edit bonus manually."
                            End If
                        End If
                    'Parking Control analysis starts here
                    Case pcLeadName 'Is Parking Control Lead
                        Errored = False
                        pc = pc + c
                        PCLead(pc) = ele.Children(2).innerText
                        For m = 0 To UBound(Managers())
                            PCLead(pc) = Replace(PCLead(pc), Managers(m), vbNullString)
                        Next m
                        PCDate(pc) = mnthName & " " & ClassDay
                        PCLoc(pc) = Split(ele.Children(3).Children(0).innerText, vbNewLine)(0)
                        c = countEmployees(PCLead(pc))
                        
                        If c > 1 Then
                            Errored = True
                            tempLead() = Split(PCLead(pc), vbNewLine)
                            For i = 0 To c - 1
                                PCLead(pc + i) = tempLead(i)
                                PCDate(pc + i) = mnthName & " " & ClassDay
                                PCe(pc + i) = True
                                PCLoc(pc + i) = ele.Children(3).Children(2).innerText
                                For Each ele2 In HTML.getElementsByClassName(dateClass)
                                    If ele2.Children.length > 1 Then
                                        'has attendants
                                        If Not ele2.Children(2).innerText = vbNullString Then 'Has attendants?
                                            'the title matches
                                            If getLocation(ele2) = getLocation(ele) Then   'NumAttendant = NumAttendant + {} because multiple PC shifts for one Lead
                                                If getLeadType(ele2) <> pcLeadName Then
                                                    NumAttendant(pc + i) = NumAttendant(pc + i) + countEmployees(ele2.Children(2).innerText)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next ele2
                                PCerr(pc) = PCerr(pc) & vbNewLine & tempLead(i)
                            Next i
                            PCerr(pc) = "Specified more than 1 Non-Steven/Brian Lead for Parking Control." & vbNewLine & vbNewLine & "Location:" & vbNewLine & PCLoc(pc) & vbNewLine & vbNewLine & "Date:" & vbNewLine & PCDate(pc) & vbNewLine & vbNewLine _
                                & "Shift Leads:" & PCerr(pc) & vbNewLine & vbNewLine & "Leads granted $15 bonus highlighted in yellow. If incorrect result, please designate ONE Lead on Shift Planning and try again OR edit bonuses manually."
                        Else 'c is not greater than 1
                            For Each ele2 In HTML.getElementsByClassName(dateClass)
                                If ele2.Children.length > 1 Then
                                    'has attendants
                                    If Not ele2.Children(2).innerText = vbNullString Then 'Has atendants?
                                        'the title matches
                                        If getLocation(ele2) = getLocation(ele) Then  'NumAttendant = NumAttendant + {} because multiple PC shifts for one Lead
                                            If getLeadType(ele2) <> pcLeadName Then
                                                NumAttendant(pc) = NumAttendant(pc) + countEmployees(ele2.Children(2).innerText)
                                            End If
                                        End If
                                    End If
                                End If
                            Next ele2
                        End If
                        If NumAttendant(pc) <= 0 Then
                            PCe(pc) = True
                            If Errored = False Then
                                PCerr(pc) = "No Non-Lead attendants found for Parking Control." & vbNewLine & vbNewLine & "Location:" & vbNewLine & PCLoc(pc) & vbNewLine & vbNewLine & "Date:" & vbNewLine & PCDate(pc) & vbNewLine & vbNewLine & "Shift Lead:" _
                                    & vbNewLine & PCLead(pc) & vbNewLine & vbNewLine & "Lead granted $15 bonus highlighted in yellow. If incorrect result, please make correction and try again OR edit bonus manually."
                            Else
                                PCerr(pc) = "No Non-Lead attendants found for Parking Control." & vbNewLine & vbNewLine & "Location:" & vbNewLine & PCLoc(pc) & vbNewLine & vbNewLine & "Date:" & vbNewLine & PCDate(pc) & vbNewLine & vbNewLine & "Shift Lead:" _
                                    & vbNewLine & PCLead(pc) & vbNewLine & vbNewLine & "Leads granted $15 bonus highlighted in yellow. If incorrect result, please make correction and try again OR edit bonus manually."
                            End If
                        End If
                    'End Case Is
                End Select
            End If
        End If
    Next ele
Next ClassDay

'WHAT THE FUCK DOES THIS DO - IT FINDS AMBIGUOUS LEAD INFO
For ClassDay = fDay To lDay
    'Check Valet and Attendant shifts
    dateClass = getClassForPayPeriod(mnth, ClassDay, yr)
    For Each ele In HTML.getElementsByClassName(dateClass)
        If ele.Children.length > 1 Then
            match = False
            'Determine if the first match has a title
            If Left(ele.Children(1).Children(0).innerText, 8) = "PP Valet" Then
                EventType = "PP"
                If ele.Children(3).Children.length > 0 Then
                    eventLocation = Split(ele.Children(3).Children(0).innerText, vbNewLine)(0)
                Else
                    eventLocation = "{Missing PP Valet Title}"
                End If
                EventDate = mnthName & " " & ClassDay
            ElseIf Left(ele.Children(1).Children(0).innerText, 12) = "PC Attendant" Then
                EventType = "PC"
                If ele.Children(3).Children.length > 0 Then
                    eventLocation = Split(ele.Children(3).Children(0).innerText, vbNewLine)(0)
                Else
                    eventLocation = "{Missing PC Attendant Title}"
                End If
                EventDate = mnthName & " " & ClassDay
            Else
                GoTo NextEle
            End If
            'Check Lead shifts
            'Loop through other events of same date (same className, which is derived from date)
            For Each ele2 In HTML.getElementsByClassName(dateClass)
                If ele2.Children.length <= 1 Then
                    GoTo NextEle2
                End If
                Select Case getLeadType(ele2)
                    Case ppLeadName
                        LeadType = "PP"
                        If ele.Children(3).Children.length > 0 Then
                            leadLocation = Split(ele.Children(3).Children(0).innerText, vbNewLine)(0)
                        Else
                            leadLocation = "{Missing PP Lead Title}"
                        End If
                        EventDate = mnthName & " " & ClassDay
                    Case pcLeadName
                        LeadType = "PC"
                        If ele.Children(3).Children.length > 0 Then
                            leadLocation = Split(ele.Children(3).Children(0).innerText, vbNewLine)(0)
                        Else
                            leadLocation = "{Missing PC Lead Title}"
                        End If
                        EventDate = mnthName & " " & ClassDay
                    Case Else
                        GoTo NextEle2
                End Select
                If eventLocation = leadLocation And EventType = LeadType Then
                    match = True
                    Exit For
                End If
NextEle2:
            Next ele2
            'Did not find a matching event
            If match = False Then
                msg = msg + 1
                If EventType = "PP" Then
                    PossibleMatch = ""
                    For Each LeadLoc In PPLoc 'Find possible matches based on first four letters or second word
                        If LeadLoc <> "" Then
                            'This was a really, really fucking retarded if statement I made...
                            If isPossibleMatch(CStr(LeadLoc), eventLocation) Then
                                PossibleMatch = PossibleMatch & vbNewLine & vbTab & LeadLoc & vbNewLine & vbTab & "Difference:" & vbNewLine & vbTab & """" & Words.WordDif(LeadLoc, eventLocation) & """, """ & Words.WordDif(eventLocation, LeadLoc) & """"
                            End If
                        End If
                    Next LeadLoc
                    If PossibleMatch <> "" Then
                        PossibleMatch = vbNewLine & vbNewLine & vbTab & "Possible Matches: " & PossibleMatch
                        UnMatchedmsg = UnMatchedmsg & vbNewLine & vbNewLine & "#" & msg & " - " & "PP" & " - " & EventDate & ": " & vbNewLine & vbTab & eventLocation & PossibleMatch
                    ElseIf PossibleMatch = "" Then
                        UnMatchedmsg = UnMatchedmsg & vbNewLine & vbNewLine & "#" & msg & " - " & "PP" & " - " & EventDate & ": " & vbNewLine & vbTab & eventLocation
                    End If
                ElseIf EventType = "PC" Then
                    PossibleMatch = ""
                    For Each LeadLoc In PCLoc 'Find possible matches based on first four letters or second word
                        If LeadLoc <> "" Then
                            If isPossibleMatch(CStr(LeadLoc), eventLocation) Then
                                PossibleMatch = PossibleMatch & vbNewLine & vbTab & LeadLoc & vbNewLine & vbTab & "Difference:" & vbNewLine & vbTab & """" & Words.WordDif(LeadLoc, eventLocation) & """, """ & Words.WordDif(eventLocation, LeadLoc) & """"
                            End If
                        End If
                    Next LeadLoc
                    If PossibleMatch <> "" Then
                        PossibleMatch = vbNewLine & vbNewLine & vbTab & "Possible Matches: " & PossibleMatch
                        UnMatchedmsg = UnMatchedmsg & vbNewLine & vbNewLine & "#" & msg & " - " & "PC" & " - " & EventDate & ": " & vbNewLine & vbTab & eventLocation & PossibleMatch
                    ElseIf PossibleMatch = "" Then
                        UnMatchedmsg = UnMatchedmsg & vbNewLine & vbNewLine & "#" & msg & " - " & "PC" & " - " & EventDate & ": " & vbNewLine & vbTab & eventLocation
                    End If
                End If
            End If
        End If
NextEle:
    Next ele
Next ClassDay
UnMatchedmsg = "Shift Leads not found for:" & vbNewLine & UnMatchedmsg

'Fill the import sheet with the bonuses
Set im = Config.getSheet_Import()
'Shift Lead Bonuses -> If V > 11 Then Bonus = 50, ElseIf V > 7 Then Bonus = 35, ElseIf V > 3 Then Bonus = 25, ElseIf V > 0 Then Bonus = 15
'Private Party
If pp = 0 Then GoTo SkipPP
For l = 1 To pp ' didn't know VBA was a base 0 language when I wrote this. Oh how young and stupid I was. Also, I was really bad at naming simple variables.
    If isManager(PPLead(l)) Then GoTo NextPP
    Select Case NumValet(l)
        Case 0 To 3
            Bonus = 15
        Case 4 To 7
            Bonus = 25
        Case 8 To 11
            Bonus = 35
        Case Is > 11
            Bonus = 50
    End Select
    For Each employee In Range(im.Cells(2, "A"), im.Cells(1, "A").End(xlDown))
        If employee = PPLead(l) And Day(employee.Offset(0, 1)) = Day(PPDate(l)) And employee.Offset(0, 3) = ppLeadName Then
            employee.Offset(0, 7) = Bonus
            If employee.Offset(0, 7).Comment Is Nothing Then
                employee.Offset(0, 7).AddComment NumValet(l) & " Valets Led"
            Else
                employee.Offset(0, 7).Comment.text text:=employee.Offset(0, 7).Comment.text & vbNewLine & NumValet(l) & " Valets Led"
            End If
            If PPe(l) = True Then
                employee.Offset(0, 7).Interior.Color = vbYellow
            End If
            Exit For
        End If
    Next employee
NextPP:
Next l
SkipPP:
'Parking Control
If pc = 0 Then GoTo SkipPC
For l = 1 To pc
    If isManager(PCLead(l)) Then GoTo NextPC
    Select Case NumAttendant(l)
        Case 0 To 3
            Bonus = 15
        Case 4 To 7
            Bonus = 25
        Case 8 To 11
            Bonus = 35
        Case Is > 11
            Bonus = 50
    End Select
    For Each employee In Range(im.Cells(2, "A"), im.Cells(1, "A").End(xlDown))
        If employee = PCLead(l) And Day(employee.Offset(0, 1)) = Day(PCDate(l)) And employee.Offset(0, 3) = pcLeadName Then
            employee.Offset(0, 7) = Bonus
            employee.Offset(0, 7).AddComment NumAttendant(l) & " Attendants Led"
            If PCe(l) = True Then
                employee.Offset(0, 7).Interior.Color = vbYellow
            End If
            Exit For
        End If
    Next employee
NextPC:
Next l
SkipPC:


endSub:

If Err <> 0 Then
    MsgBox Err.Number & ": " & Err.Description
End If

End Sub

Private Function getClassForPayPeriod(month_ As Integer, day_ As Integer, year_ As Integer) As String 'how the className is formatted for all dates we want to analyze

    Dim month_value As String
    Dim day_value As String
    
    month_value = IIf(month_ < 10, "0" & month_, CStr(month_))
    day_value = IIf(day_ < 10, "0" & day_, CStr(day_))

    getClassForPayPeriod = "tl_" & month_value & "_" & day_value & "_" & year_
End Function

Private Function getLocation(ele_ As IHTMLElement) As String
    
    Dim lastChild As IHTMLElement
    Dim grandChild As IHTMLElement
    Dim text As String
    Dim tokens() As String
    Dim finalText As String
    
    Set lastChild = ele_.Children(3)
    Set grandChild = lastChild.Children(0)
    text = grandChild.innerText
    tokens() = Strings.Split(text, " - ")
    text = tokens(0)
    tokens() = Strings.Split(text, "Private")
    text = tokens(0)
    tokens() = Strings.Split(text, "Parking")
    text = tokens(0)
    
    finalText = UCase(text)
    
    getLocation = finalText
End Function

Private Function countEmployees(employeeList As String) As Integer
'one employee has a space in their name
'each space means 1 employee
'it is very important that employees don't have their middle name in there

countEmployees = Words.CharInst(employeeList, " ")

End Function

Private Function isManager(employeeName As String) As Boolean

    isManager = employeeName = "Bauer, Brian" Or employeeName = "Bauer, Steven"

End Function

Private Function isPossibleMatch(leadLocation As String, eventLocation As String) As Boolean

    Dim leadLocationStart As String, eventLocationStart As String
    Dim partialLeadLocation As String, partialEventLocation As String
    
    leadLocationStart = Left(leadLocation, 4)
    eventLocationStart = Left(eventLocation, 4)
    partialLeadLocation = Mid(leadLocation, InStr(1, leadLocation, " ") + 1, IIf(InStr(InStr(1, leadLocation, " ") + 1, leadLocation, " ") - 1 < 0, 0, InStr(InStr(1, leadLocation, " ") + 1, leadLocation, " ") - 1))
    partialEventLocation = Mid(eventLocation, InStr(1, eventLocation, " ") + 1, IIf(InStr(InStr(1, eventLocation, " ") + 1, eventLocation, " ") - 1 < 0, 0, InStr(InStr(1, eventLocation, " ") + 1, eventLocation, " ") - 1))

    isPossibleMatch = UCase(leadLocationStart) = UCase(eventLocationStart) Or UCase(partialLeadLocation) = UCase(partialEventLocation)
    
End Function

Private Function getLeadType(ele_ As IHTMLElement) As String
    
    getLeadType = Left(ele_.Children(1).Children(0).innerText, 7)
    
End Function

