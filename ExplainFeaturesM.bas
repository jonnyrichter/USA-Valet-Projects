Attribute VB_Name = "ExplainFeaturesM"
Option Explicit

Sub ExplainFeatures()

Dim CtrlJ As String
Dim CtrlShiftS As String
Dim SheetDblClick As String
Dim CtrlR As String
Dim CtrlN As String

CtrlJ = "1. Ctrl + J - Use on any sheet to go to the next open slot. Useful shortcut when filling out time cards manually." & vbNewLine & vbNewLine
CtrlShiftS = "2. Press Ctrl + Shift + S - Sort names on all sheets (including 'Total') alphabetically and re-sizes columns." & vbNewLine & vbNewLine
SheetDblClick = "3. Double Click - When there is only one empty slot left on a sheet (top right, with a used slot before it), use on empty slot to create a new column." & vbNewLine & vbNewLine
CtrlR = "4. (This is currently broken - do not use) Ctrl + R - On the Private Party (PP) or Parking Control (PC) Sheets, while selecting the ""Reimbursement"" row under the employee you wish to reimburse for travel will extract from Google Maps their minutes traveled without traffic." & vbNewLine & vbNewLine
CtrlN = "5. Ctrl + N - bring up a mini menu of all sheets to easily navigate through them" & vbNewLine & vbNewLine

MsgBox CtrlJ & CtrlShiftS & SheetDblClick & CtrlR & CtrlN, vbInformation, "Special Button and Click Commands"
End Sub

