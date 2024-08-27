Attribute VB_Name = "ChooseSheetM"
Option Explicit

Sub ChooseSheet()
Attribute ChooseSheet.VB_ProcData.VB_Invoke_Func = "n\n14"
    Dim ws As Worksheet
     
    Application.CommandBars("Workbook Tabs").ShowPopup
    Set ws = ActiveSheet
End Sub
