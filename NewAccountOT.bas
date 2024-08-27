Attribute VB_Name = "NewAccountOT"
Option Explicit

Private Const getClass As String = "NewAccountOT"

Private Const testSheetName = "TR Lead"
Private Const testCopySheet = "RC"
Private Const sheetToInsertBefore = "MT" '"T.5&2T Hours" <= use this (t.5&2t) for lead type accounts
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub createNewRowTest() ' test sub - add row
    Call System.Update(False)
    Call createNewRow(testSheetName, testCopySheet)
    Call System.Update(True)
End Sub

Public Sub deleteSheetOnOTTest() ' test sub - delete
    System.Update False
    deleteOnOT testSheetName
    System.Update True
End Sub

Public Sub copyNewColorToNewAccountFromAccountSheet()
    Const sheetName = testSheetName 'this is redundant but whatever
    Dim ws As Worksheet
    Dim fillColor As Long, fontColor As Long
    Dim c As Range
    
    Set ws = Config.getSheet(sheetName)
    
    fillColor = ws.Range("A1").Interior.Color
    fontColor = ws.Range("A1").Font.Color
    
    System.unprotectSheet Config.getSheet_OT()
    
    For Each c In Config.getSheet_OT().Range("B:B")
        If (c.value = sheetName) Then
            c.Interior.Color = fillColor
            c.Font.Color = fontColor
            Exit For
        End If
    Next c
    
    Call copyNewColorToNewAccount
    
    System.protectSheet Config.getSheet_OT()
    
End Sub
Public Sub copyNewColorToNewAccount() ' do after createNewRowTest(), once you set the correct color to the first instance
    Dim c As Range
    Const sheetName As String = testSheetName
    Dim foundFirst As Boolean
    Dim i As Integer
    
    For Each c In Config.getSheet_OT().Range("B:B")
        If (c.value = sheetName) Then
            i = i + 1
            If Not foundFirst Then
                c.Copy
                foundFirst = True
            Else
                c.PasteSpecial xlPasteFormats
            End If
        End If
        If i = 16 Then Exit For
    Next c
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub createNewRow(newSheetAbbreviation As String, copyFromName As String)
    Call insertOnOT '()
    copyFromRowOnOT newSheetAbbreviation, copyFromName
    
End Sub


Public Sub copyFromRowOnOT(newName As String, copyFrom As String)

    Dim ot As Worksheet
    Dim i As Long, r As Long
    Dim rangeToCopy As String, rangeToPaste As String
    Dim rangesToCopy() As String, rangesToPaste() As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set ot = Config.getSheet_OT()
    System.unprotectSheet ot
    
    rangeToCopy = getRange(copyFrom) ' get the range of the new sheetName
    rangeToPaste = getBlankRange() ' get the range of the blank cells
    
    rangesToCopy = Strings.Split(rangeToCopy, ",") ' separate the ranges into (15 or 16) pieces
    rangesToPaste() = Strings.Split(rangeToPaste, ",")
    
    For i = 0 To UBound(rangesToPaste())
        r = CInt(Strings.Replace(Strings.Split(rangesToPaste(i), ":")(0), "A", vbNullString)) ' get 12 from A12:BF12,A29:BF29 etc...
        ot.Range(rangesToCopy(i)).Copy
        ot.Range(rangesToPaste(i)).PasteSpecial ' paste should format to new row
        ot.Range(rangesToPaste(i)).Replace "'" & copyFrom & "'!", "'" & newName & "'!" ' change out for parameter, as literal
        ot.Range(rangesToPaste(i)).Replace copyFrom & "!", newName & "!" ' change out for parameter
        ot.Range(rangesToPaste(i)).Replace copyFrom & "Emp", newName & "Emp" ' change out for EMP ranges
        ot.Range("B" & r).value = newName ' put the new account abbreviation in B column
    Next i
    
    Application.CutCopyMode = False
    System.unprotectSheet ot
    
End Sub


Public Sub insertOnOT() ' inserts a blank row before every Training/Management row

    System.unprotectSheet Config.getSheet_OT()
    
    Config.getSheet_OT().Range(getRange(sheetToInsertBefore)).Insert xlShiftDown
    
    System.protectSheet Config.getSheet_OT()

End Sub

Public Sub deleteBlankRowOnOT() ' removes all the blank rows created by insertOnOt()

    deleteOnOT vbNullString
    
End Sub

Public Sub deleteOnOT(accountAbbreviation As String) ' removes all the rows with given sheet name

    System.unprotectSheet Config.getSheet_OT()
    
    Config.getSheet_OT().Range(getRange(accountAbbreviation)).Delete xlShiftUp
    
    System.protectSheet Config.getSheet_OT()
    
End Sub

Public Function getBlankRange() As String ' returns the entire range of every blank row within OT

    getBlankRange = getRange(vbNullString)
    
End Function


Public Function getRange(sheetName As String) As String ' gets the start and end column of OT sheet with every row that equals sheet name (vbNullString for blank rows)

    Dim ot As Worksheet
    Dim r As Long, j As Long, lastRow As Long, startRow As Long, rowDifference As Long
    Dim rowRange As String, comma As String, lastCol As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    log.setClass(getClass).setMethod ("getRange")
    Set ot = Config.getSheet_OT()
    
    lastRow = getLastRowInCol(ot, "B") 'the column with the account abbreviations
    lastCol = getLastColInRow(ot, 2) 'Gets the column name of the last used column
    
    For j = 1 To lastRow
        If ot.Range("B" & j).value = sheetName Then 'Training/Manager Meeting Sheet Name
            If startRow > 0 Then  'cause it'll be zero when it's analyzing itself
                rowDifference = j - startRow 'on second match the difference between start row and next occurence is found
                Exit For
            End If
            If startRow = 0 Then startRow = j 'on first match the start row is that row
        End If
    Next j
    
    If j = lastRow Then
        log.error "iteration made it to last row"
        System.Update True
        End
    End If
    
    For r = startRow To lastRow Step rowDifference 'starting row to delete at, ending row,
        rowRange = rowRange & comma & "A" & r & ":" & lastCol & r
        comma = ","
    Next r
    
    getRange = rowRange
    
End Function

Public Function getLastRowInCol(s As Worksheet, colName As String) As Long ' given a sheet and column name, returns its last row
    getLastRowInCol = s.Range(colName & s.Rows.count).End(xlUp).Row
End Function
Public Function getLastColInRow(s As Worksheet, rowNum As Long) As String ' given a sheet and row number, returns its column by name
    getLastColInRow = Words.col(s.Range(Words.col(s.Columns.count) & rowNum).End(xlToLeft).Column)
End Function
