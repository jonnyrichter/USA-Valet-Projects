Attribute VB_Name = "FitSheetM"
Option Explicit

Sub FitSheet()

    System.unprotectSheet ThisWorkbook.ActiveSheet
    ThisWorkbook.ActiveSheet.Cells.EntireColumn.AutoFit
    System.protectSheet ThisWorkbook.ActiveSheet

End Sub
