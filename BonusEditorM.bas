Attribute VB_Name = "BonusEditorm"
Option Explicit
Option Compare Text

Sub EditBonus()

Dim eRow As Variant
Dim Bonus As Variant, newBonus As Variant
Dim Cell As Variant
Dim Colored As Double
Dim im As Worksheet

On Error GoTo endSub
Set im = Config.getSheet_Import()

If im.Cells(2, 1) = "" Then
    Beep
    MsgBox "Time cards have not been imported yet", , "Cannot Edit"
    Exit Sub
End If

frmPword.Show

System.unprotectSheet im

For Each Cell In im.Range("A:A")
    If Cell = vbNullString Then Exit For
    If Cell.Row = 1 Then GoTo Skip
    Cell.Offset(0, 8) = Cell.Row
Skip:
Next Cell

tryAgain:

eRow = InputBox("Which row contains the bonus you would like to edit?", "Choose Row", "2")
If eRow = "" Then GoTo endSub
If eRow = 1 Or CInt(eRow) > im.Cells(1, 1).End(xlDown).Row Then
    MsgBox "Please enter in a number from 2 to " & im.Cells(1, 1).End(xlDown).Row, vbOKOnly, "Invalid Selection"
    GoTo tryAgain
End If

eRow = CInt(eRow)
Colored = im.Cells(eRow, "H").Interior.Color
Range(im.Cells(eRow, "A"), im.Cells(eRow, "H")).Select
im.Cells(eRow, "H").Interior.Color = vbCyan

TryAgain2:

reBonus:
Bonus = InputBox("By how much would you like to change this bonus?", "Enter Dollar Amount", "0")
If IsNumeric(Bonus) = False Then
    If Bonus = "" Then GoTo endSub
    MsgBox "Please choose a positive or negative amount to alter the bonus by.", , "Error"
    GoTo reBonus
End If
If Bonus = "" Then
    im.Cells(eRow, "H").Interior.Color = Colored
    im.Range("I:I").ClearContents
    System.protectSheet im
    Exit Sub
End If
If IsNumeric(Bonus) = False Then MsgBox "Please enter a valid dollar number", , "Invalid": GoTo TryAgain2

Bonus = CDbl(Bonus)
If Bonus < 0 Then im.Cells(eRow, "H").Interior.Color = vbRed
If Bonus = "" Or Bonus = 0 Then im.Cells(eRow, "H").Interior.Color = Colored
newBonus = im.Cells(eRow, "H") + Bonus

'verify the amount is correct
If MsgBox("Bonus will be changed from $" & im.Cells(eRow, "H") & " to $" & newBonus & ". Is this correct?", vbYesNo, "Please Verify") = vbNo Then
    GoTo reBonus
End If

im.Cells(eRow, "H") = newBonus
im.Cells.EntireColumn.AutoFit

endSub:

If Err = 13 Then
    Err.clear
    MsgBox "Please enter in a number from 2 to " & im.Cells(1, 1).End(xlDown).Row, vbOKOnly, "Invalid Selection"
    GoTo tryAgain
ElseIf Err <> 0 Then
    MsgBox "Error #: " & Err.Number & " - " & Err.Description, vbCritical
End If

im.Range("I:I").ClearContents
System.protectSheet im

End Sub

Sub RemoveBonus()

Dim eRow As Variant
Dim YorN As Integer
Dim Cell As Variant
Dim Colored As Double
Dim nShort As Range
Dim nLong As Range
Dim im As Worksheet

On Error GoTo endSub
Set im = Config.getSheet_Import()
Set nShort = Ranges.getEmployeeListRange()
Set nLong = Ranges.getTotalEmpRange()

If im.Cells(2, 1) = "" Then
    Beep
    MsgBox "Time cards have not been imported yet", , "Cannot Remove"
    Exit Sub
End If

frmPword.Show

System.unprotectSheet im

For Each Cell In im.Range("A:A")
    If Cell = "" Then Exit For
    If Cell.Row = 1 Then GoTo Skip
    Cell.Offset(0, 8) = Cell.Row
Skip:
Next Cell

tryAgain:

eRow = InputBox("Which row contains the bonus you would like to remove?", "Choose Row", "2")
If eRow = "" Then GoTo endSub
If eRow = 1 Or CInt(eRow) > im.Cells(1, 1).End(xlDown).Row Then
    MsgBox "Please enter in a number from 2 to " & im.Cells(1, 1).End(xlDown).Row, vbOKOnly, "Invalid Selection"
    GoTo tryAgain
End If

eRow = CInt(eRow)
Colored = im.Cells(eRow, "H").Interior.Color
Range(im.Cells(eRow, "A"), im.Cells(eRow, "H")).Select
im.Cells(eRow, "H").Interior.Color = vbRed

If im.Cells(eRow, "H") >= 0 Then
    YorN = MsgBox("Are you sure you would like to remove $" & im.Cells(eRow, "H") & " bonus for: " & _
        vbNewLine & vbNewLine & im.Cells(eRow, "A") & vbNewLine & "On " & _
        im.Cells(eRow, "B") & vbNewLine & "As a " & im.Cells(eRow, "D") & "?", vbYesNo, "Remove?")
Else
    YorN = MsgBox("Are you sure you would like to remove ($" & Abs(im.Cells(eRow, "H")) & ") bonus for: " & _
        vbNewLine & vbNewLine & im.Cells(eRow, "A") & vbNewLine & "On " & _
        im.Cells(eRow, "B") & vbNewLine & "As a " & im.Cells(eRow, "D") & "?", vbYesNo, "Remove?")
End If
If YorN = vbYes Then
    im.Cells(eRow, "H").ClearContents
Else
    im.Cells(eRow, "H").Interior.Color = Colored
End If

im.Cells.EntireColumn.AutoFit

endSub:

If Err = 13 Then
    Err.clear
    MsgBox "Please enter in a number from 2 to " & im.Cells(1, 1).End(xlDown).Row, vbOKOnly, "Invalid Selection"
    GoTo tryAgain
ElseIf Err <> 0 Then
    MsgBox "Error #: " & Err.Number & " - " & Err.Description, vbCritical
End If

im.Range("I:I").ClearContents
System.protectSheet im

End Sub

