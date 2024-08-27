Attribute VB_Name = "SortSMS"
Option Explicit

Public Sub SortByName()

SortBy 1

End Sub

Public Sub SortByDate()

SortBy 2

End Sub

Public Sub SortByLocation()

SortBy 3

End Sub

Public Sub SortByStartDif()

SortBy 9, False

End Sub

Public Sub SortByEndDif()

SortBy 10, False

End Sub

Private Sub SortBy(SortColumn As Integer, Optional Asc As Boolean = True)
'Sort Fields
Dim lastRow As Integer
Dim sms As Worksheet
Set sms = Config.getSheet_SMS()

If isEmpty(sms.Cells(2, 1)) Then
    MsgBox "Data not yet imported", vbCritical, "Too Soon, Breh"
    End
End If

System.Update False

System.unprotectSheet sms

lastRow = sms.Cells(1, 1).End(xlDown).Row

sms.Sort.SortFields.clear
sms.Sort.SortFields.add key:=Range(sms.Cells(2, SortColumn), sms.Cells(lastRow, SortColumn)) _
    , SortOn:=xlSortOnValues, Order:=IIf(Asc, xlAscending, xlDescending), DataOption:=xlSortNormal
With sms.Sort
    .SetRange Range(sms.Cells(2, 1), sms.Cells(lastRow, IIf(Asc, 8, 10)))
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

System.protectSheet sms
System.Update True

End Sub

