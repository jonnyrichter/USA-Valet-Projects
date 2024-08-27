Attribute VB_Name = "SortSheetsM"
Option Explicit

Public Sub SortSheets() 'Separate Total and other sheets into TWO subs that are called from this one.
Attribute SortSheets.VB_Description = "Sorts all sheets"
Attribute SortSheets.VB_ProcData.VB_Invoke_Func = " \n14"

Dim ws As Worksheet
Dim WsName As String
Dim yesNo As Boolean

On Error GoTo endSub


yesNo = MsgBox("Alphabetize all names in this workbook and refit the cells to the correct size?", vbYesNo, "Continue?") = vbYes
If Not yesNo Then End

ThisWorkbook.Unprotect Passwords.getDevPassword()
System.Update False

For Each ws In ThisWorkbook.Worksheets

    System.unprotectSheet ws
    If ws.name = Config.getTotalSheetName() Then
        ws.Sort.SortFields.clear
        ws.Sort.SortFields.add key:=Ranges.getTotalEmpRange(), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
        With ws.Sort
            .SetRange Ranges.getTotalSortFieldRange()
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    
        ws.Cells.EntireColumn.AutoFit
        GoTo NextWS

    ElseIf PayPeriodTypes.isMonthlySheet(ws.name) Or PayPeriodTypes.isSemimonthlySheet(ws.name) Then
        WsName = ws.name
        
        ws.Sort.SortFields.clear
        ws.Sort.SortFields.add key:=Range(WsName & "Emp"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
        With ws.Sort
            .SetRange Ranges.getSortFieldRange(WsName)
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlLeftToRight
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
NextWS:
ws.Cells.EntireColumn.AutoFit
System.protectSheet ws
Next ws

endSub:
If Err <> 0 Then
    MsgBox "Error #: " & Err.Number & " - " & Err.Description, vbCritical
End If

Config.getSheet_Total().Activate

System.Update True
ThisWorkbook.Protect Passwords.getDevPassword()

Beep

End Sub

