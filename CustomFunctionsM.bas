Attribute VB_Name = "CustomFunctionsm"
Option Explicit
'Custom Functions

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function HRS(StartTime As Double, EndTime As Double) As Double

On Error GoTo endSub

'Finds the length of a single employees shift
If EndTime < StartTime Then
    HRS = (EndTime + 1 - StartTime) * 24
Else
    HRS = (EndTime - StartTime) * 24
End If

endSub:
If Err <> 0 Then
    MsgBox "Error #: " & Err.Number & " - " & Err.Description, vbCritical
End If

End Function

Public Function SHIFT(StartTimeRange As Range, endTimeRange As Range, LotBlockerPreArrival As Double) As Double

On Error GoTo endSub

'Finds the length of the shift regardless of going past midnight (subtracts time that lot blocker was there before first valet)

Dim s As Double, e As Double
Dim t As Variant

s = WorksheetFunction.Min(StartTimeRange)

If WorksheetFunction.CountA(endTimeRange) = 0 Then
    e = 0
    s = 0
ElseIf WorksheetFunction.Min(endTimeRange) < 0.25 Then
    e = 0
    For Each t In endTimeRange
        If t.value < 0.25 And t.value > e Then
            e = t.value
        Else
            e = e
        End If
    Next
    e = e + 1
Else
    e = WorksheetFunction.Max(endTimeRange)
End If

SHIFT = (e - s) * 24 + LotBlockerPreArrival

endSub:
If Err <> 0 Then
    MsgBox "Error #: " & Err.Number & " - " & Err.Description, vbCritical
End If

End Function

