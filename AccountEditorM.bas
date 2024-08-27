Attribute VB_Name = "AccountEditorM"
Option Explicit

Public Sub editAccountName()

    Dim oldTabName As String, oldAccountName As String
    Dim newTabName As String, newAccountName As String
    Dim wb As Workbook
    Dim rangeNames() As String
    Dim rangeName As name
    Dim n As Long
    Dim editWs As Worksheet
    Dim tot As Worksheet, ot As Worksheet, im As Worksheet
    Dim totEditRange As Range, otEditRange As Range, imEditRange As Range
    Dim totCell As Range, otCell As Range, imCell As Range
    Dim count As Long
    Const maxCount As Long = 16 'no payperiod will have more than 16 days - OT has exactly 16 days

    On Error GoTo error_handler
    
    frmPword.Show
    
    'MsgBox "This method is still beta. Please save any unsaved work before proceeding.", vbExclamation, "Warning!"
    
    frmEditAccountName.Show
    
    System.Update False

    oldTabName = editAccountDataStore.getOldTabName()
    oldAccountName = editAccountDataStore.getOldFullName()
    newTabName = editAccountDataStore.getNewTabName()
    newAccountName = editAccountDataStore.getNewFullName()
    
    Set wb = Config.getWorkbook()
    Set tot = Config.getSheet_Total()
    Set ot = Config.getSheet_OT()
    Set im = Config.getSheet_Import()
    Set editWs = Config.getSheet(oldTabName)
    
    System.unprotectWorkbook
    System.unprotectSheet editWs
    System.unprotectSheet tot
    System.unprotectSheet ot
    System.unprotectSheet im
    
    editWs.name = newTabName
    editWs.Range("A1").value = newAccountName
    
    'Edit "Total" sheet
    For Each totCell In tot.Range("1:1")
    
        If totCell.value = oldTabName Then
            
            totCell.value = newTabName
            Exit For 'only has 1 instance
            
        End If
            
    Next totCell
    
    'Edit "Import" sheet
    For Each imCell In im.Range("1:1")
    
        If imCell.value = oldTabName Then
        
            imCell.value = newTabName
            imCell.Offset(1).Formula = Strings.Replace(imCell.Offset(1).Formula, """" & oldAccountName & """", """" & newAccountName & """")
            Exit For 'only has 1 instance
        
        End If
        
    Next imCell
    
    'Edit "OverTime" sheet
    count = 0
    For Each otCell In ot.Range("B:B")
    
        If otCell.value = oldTabName Then
        
            otCell.value = newTabName
            count = count + 1
            If count = maxCount Then
                Exit For 'only fix 16 times
            End If
            
        End If
        
    Next otCell
    
    'Fix the Range names because those are called later in many other routines
    rangeNames() = Ranges.getRangeNames(oldTabName)
    
    For n = 0 To UBound(rangeNames())
    
        Set rangeName = wb.Names(rangeNames(n))
        
        rangeName.name = Strings.Replace(rangeNames(n), oldTabName, newTabName, , 1)
        
    Next n
    
    'Set every thing back to user-safe mode
    System.protectWorkbook
    System.protectSheet editWs
    System.protectSheet tot
    System.protectSheet ot
    System.protectSheet im
    
    System.Update True
    
    editWs.Activate
    
Exit Sub

error_handler:

    MsgBox "There was an unexpected error. Please close the workbook WITHOUT saving.", vbCritical, "Error!"

    System.protectWorkbook
    System.protectSheet editWs
    System.protectSheet tot
    System.protectSheet ot
    System.protectSheet im

    System.Update True
End Sub
