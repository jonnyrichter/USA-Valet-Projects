Attribute VB_Name = "UnprotectAllSheetsm"
Option Explicit

Sub UnprotectAllSheets()

    Dim ws As Worksheet
    Dim wb As Workbook
    
    Set wb = ThisWorkbook
    
    System.unprotectWorkbook
    
    For Each ws In wb.Worksheets
        System.unprotectSheet ws
    Next ws

End Sub

Sub ProtectAllSheets()

    Dim ws As Worksheet
    Dim wb As Workbook
    
    Set wb = ThisWorkbook
    
    System.protectWorkbook
    
    For Each ws In wb.Worksheets
        System.protectSheet ws
    Next ws

End Sub

