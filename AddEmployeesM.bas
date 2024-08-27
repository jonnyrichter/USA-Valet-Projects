Attribute VB_Name = "AddEmployeesm"
Option Explicit
Option Compare Text
Private ot As Worksheet, total As Worksheet

Sub AddEmployees()   'Called from ImportFromSP()

Dim import As Worksheet
Dim tName As Range, iName As Range
Dim temp As Range, iEmp As Range
Dim n As Integer
Dim match As Boolean
Dim fixedName As String, f() As String

Set import = Config.getSheet_Import()
Set total = Config.getSheet_Total()
Set ot = Config.getSheet_OT()
Set iEmp = Range(import.Cells(2, 1), import.Cells(2, 1).End(xlDown)) 'From A2 to the last employee on "Import"
Set temp = Ranges.getTotalEmpRange()

On Error GoTo endSub
'Add names from Import not in 'Total'
For Each iName In iEmp 'Import.Range(A:A)

    f() = Split(iName.value, " ") 'Marlin Skogberg -> Skogberg, Marlin
    fixedName = f(1) & ", " & f(0)
    
    For n = 1 To iName.Row - 1
        If fixedName = total.Cells(n, 1) Then GoTo NextiName
    Next n
    match = False
    For Each tName In temp
        If fixedName = tName.value Then
            match = True
            Exit For
        End If
    Next tName
    If match = False Then
        If total.Cells(1, 1).End(xlDown) = "Total" Then 'Make sure there aren't more employees than rows by seeing if the .End method returns "Total" row as last one left
            Call addRowToTotalEmployees
        End If
        total.Cells(1, 1).End(xlDown).Offset(1, 0) = fixedName
    End If
NextiName:
Next iName

'Remove names from 'Total' Not in Time Cards
For Each tName In temp
    If tName.value = vbNullString Then GoTo FoundMatch
    match = False
    
    f() = Split(tName.value, ", ") 'Skogberg, Marlin -> Marlin Skogberg
    fixedName = f(1) & " " & f(0)
    
    For Each iName In iEmp 'Match to Time Cards
        If fixedName = iName.value Then
            match = True
            GoTo FoundMatch
        End If
    Next iName
    If match = False Then
        total.Cells(tName.Row, "B").ClearContents
        total.Cells(tName.Row, "C").ClearContents
        tName.ClearContents
    End If
FoundMatch:
Next tName

endSub:
System.getError Err
System.displayError

End Sub

Public Sub addRowToTotalEmployees()
Dim c As Integer, rTH As Integer, rOT As Integer
Dim shortHandCol As String


MsgBox "This feature is broken and has been disabled. Please contact Jon Richter to either add rows or enable this feature", vbCritical, "Feature Disabled"
End

On Error GoTo endSub
System.Update False

If total Is Nothing Then
    Set total = Config.getSheet_Total()
    Set ot = Config.getSheet_OT()
End If
    c = total.Cells(1, 1).End(xlToRight).End(xlToRight).Column 'Column of the "Shorthand"
    shortHandCol = Words.col(CLng(c))
    System.unprotectSheet total
    System.unprotectSheet ot
    
    'Add row to 'Total'
    If total.Cells(1, 1).End(xlDown).value = "Total" Then
        rTH = total.Cells(1, 1).End(xlDown).Row - 1 'Row of last employee before category total - if there is no empty slots left
    Else
        rTH = total.Cells(1, 1).End(xlDown).End(xlDown).Row - 1 'Row of last employee before category total - if called from button
    End If
    Range(total.Cells(rTH, 1), total.Cells(rTH, c)).Insert xlShiftDown, xlFormatFromLeftOrAbove 'Insert before the previously mentioned row
    Range(total.Cells(rTH - 1, 1), total.Cells(rTH - 1, c)).AutoFill Range(total.Cells(rTH - 1, 1), total.Cells(rTH, c)), xlFillDefault 'Copy the formulas from previous row
    Range(total.Cells(rTH, "A"), total.Cells(rTH, "C")).ClearContents 'Clear the name and wage rate
    
    'Add column to 'OT' (It's automatic??? - No)
    rOT = ot.Cells(1, 1).End(xlDown).End(xlDown).End(xlDown).End(xlDown).Row '"Total OT owed" row - fixed
    c = ot.Cells(1, 1).End(xlToRight).Column 'Last column in "OT" sheet
    Range(ot.Cells(1, c), ot.Cells(rOT, c)).Insert xlShiftToRight, xlFormatFromLeftOrAbove 'Add Column to OT
    Range(ot.Cells(2, c - 1), ot.Cells(rOT, c - 1)).AutoFill Range(ot.Cells(2, c - 1), ot.Cells(rOT, c)), xlFillDefault 'Copy the previous formulas
    ot.Cells(1, c) = "=Total!$" & shortHandCol & rTH 'Make it reference the correct row from "Total"
    ot.Cells.Columns.AutoFit
    
    System.protectSheet ot
    System.protectSheet total

endSub:
System.Update True
System.getError Err
System.displayError

End Sub
