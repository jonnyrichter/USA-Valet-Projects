Attribute VB_Name = "Module1"
Option Explicit

'todo make this fix itself based on total - for now I have to add the columns and then fix the formulas
Public Sub fixOTRow1Numbers()

    Dim ot As Worksheet
    Dim row1 As Range
    Dim c As Range
    Dim i As Long
    Dim totalEmpColString As String, temp As String
    Dim lastCol As Boolean
    Dim s As Long
    
    Dim lineStyleValue As Long
    
    Set ot = Config.getSheet_OT()
    Set row1 = ot.Range("1:1")
    
    System.unprotectSheet ot
    
    lastCol = False
    totalEmpColString = row1(1, 3).Formula
    totalEmpColString = Split(totalEmpColString, "$")(1)
    temp = Left(totalEmpColString, 2)
    s = CLng(Replace(totalEmpColString, temp, ""))
    totalEmpColString = temp
    
    For Each c In row1
        
    
        If (c.Column > 2) Then
            If c.Borders(xlEdgeRight).LineStyle > 0 Then
                lastCol = True
            End If
        
            c.Formula = "=Total!$" & totalEmpColString & s
            s = s + 1
        End If
        
        If lastCol Then Exit For
    Next c
    
    System.protectSheet ot
End Sub

Public Sub listSheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.name
    Next ws
End Sub

