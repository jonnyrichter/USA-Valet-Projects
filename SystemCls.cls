VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "SystemCls"

Public errorValue As Integer
Public errorDescription As String
Public Sub getError(errorObj As ErrObject)
errorValue = errorObj.Number
errorDescription = errorObj.Description
End Sub

Public Sub Update(enable As Boolean)

    Application.DisplayAlerts = enable
    Application.Cursor = IIf(enable, xlDefault, xlWait)
    Application.Calculation = IIf(enable Or VersionControl.TestStatus, xlCalculationAutomatic, xlCalculationManual)
    Application.ScreenUpdating = enable
    ActiveWindow.DisplayHorizontalScrollBar = enable
    ActiveWindow.DisplayVerticalScrollBar = enable

End Sub

Public Sub displayError()
On Error Resume Next
If errorValue <> 0 Then
    System.Update True
    MsgBox "Err: " & errorValue & vbNewLine & errorDescription, vbCritical
    End
End If
End Sub

Public Sub unprotectSheet(s As Worksheet)
    s.Unprotect Passwords.getDevPassword()
End Sub
Public Sub protectSheet(s As Worksheet)
    s.Protect Passwords.getDevPassword()
End Sub
Public Sub unprotectWorkbook()
    ThisWorkbook.Unprotect Passwords.getDevPassword()
End Sub
Public Sub protectWorkbook()
    ThisWorkbook.Protect Passwords.getDevPassword()
End Sub



