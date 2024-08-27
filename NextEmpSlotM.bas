Attribute VB_Name = "NextEmpSlotm"
Option Explicit

Sub NextEmpSlot()
Attribute NextEmpSlot.VB_Description = "Moves to the next empty slot of the top row on FH or T22."
Attribute NextEmpSlot.VB_ProcData.VB_Invoke_Func = " \n14"

On Error GoTo endSub

Cells(1, 1).End(xlToRight).Offset(0, 1).Select

endSub:
If Err <> 0 Then
    MsgBox "Error #: " & Err.Number & " - " & Err.Description, vbCritical
End If

End Sub
