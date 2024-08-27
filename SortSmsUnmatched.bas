Attribute VB_Name = "SortSmsUnmatched"
Option Explicit

Public Sub SortUnmatchedByName()

SortByUnmatched "M"

End Sub

Public Sub SortUnmatchedByDate()

SortByUnmatched "N"

End Sub

Public Sub SortUnmatchedByLocation()

SortByUnmatched "O"

End Sub

Public Sub SortUnmatchedByClockTime()

SortByUnmatched "P"

End Sub

Public Sub SortUnmatchedByClockType()

SortByUnmatched "Q"

End Sub

Private Sub SortByUnmatched(SortColumn As String)
'Sort Fields
Dim lastRow As Integer
Dim sms As Worksheet
Set sms = Config.getSheet_SMS()

If isEmpty(sms.Cells(2, "M")) Then
    MsgBox "No unmatched clock time to sort", vbCritical, "Can't do it, Breh!"
    End
End If

System.Update False

System.unprotectSheet sms

lastRow = sms.Cells(1, "M").End(xlDown).Row

sms.Sort.SortFields.clear
sms.Sort.SortFields.add key:=Range(sms.Cells(2, SortColumn), sms.Cells(lastRow, SortColumn)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With sms.Sort
    .SetRange Range(sms.Cells(2, "M"), sms.Cells(lastRow, "Q"))
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

System.protectSheet sms

System.Update True

End Sub



