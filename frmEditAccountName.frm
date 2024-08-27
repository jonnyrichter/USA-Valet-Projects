VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditAccountName 
   Caption         =   "Edit Account Name"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "frmEditAccountName.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditAccountName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const this As String = "frmEditAccountName"
Private oldTabName, oldAccountName As String
Private newTabName, newAccountName As String

'This form called from AccountEditorM.editAccountName() and then data is passed to that Sub via EditAccountNameDataStoreCls.{set|get}
'btnSubmit_Click() <> verifyDataOnSubmit() <> errorMsg() <> hideAndReload() </> </> <> weirdDataMsg() <> hideAndReload() </> </> </> <> confirmUserInput() <> hideAndReload() </> </>
'btnCancel_Click() <> cancelOperation() </>
'UserForm_Initialize() <> populateComboBox() </>
'populateCombobox()
'UserForm_QueryClose()

'@Info (Click Submit - Initiates process with validation)
Private Sub btnSubmit_Click()

    Dim ws As Worksheet
    Dim cellA1 As Range
    
    'Store chosen values in variables for debugging
    oldAccountName = Me.cmbAccountToEdit.text
    newAccountName = Me.txtNewName.value
    newTabName = Me.txtNewAbbrv.value
    
    'Send those values to the instance of GlobalVariables.EditAccountNameDataStoreCls
    editAccountDataStore.setOldFullName (oldAccountName)
    editAccountDataStore.setNewFullName (newAccountName)
    editAccountDataStore.setNewTabName (newTabName)
    
    
    For Each ws In ThisWorkbook.Worksheets 'loop through sheets to get the tabName based on Account Name in A1
        
        Set cellA1 = ws.Range("A1")
        
        If cellA1.value = oldAccountName Then
        
            oldTabName = ws.name 'Store the tab name
            Exit For
        
        End If
        
    Next ws
    
    'add error checking, separate sub
    Call verifyDataOnSubmit
    
    'Make sure user is sure about their choices
    Call confirmUserInput
    
    editAccountDataStore.setOldTabName (oldTabName)
    
    Unload Me
    
End Sub

'@Info (validates that the user is sure of their input)
Private Sub confirmUserInput()

    Dim message As String
    Dim saidYes As Boolean
    
    message = "You've chosen to rename:" & Words.vLine(2) & _
        "Sheet - " & oldTabName & vbNewLine & "Account - " & oldAccountName & Words.vLine(2) & _
        " To:" & Words.vLine(2) & _
        "Sheet - " & newTabName & vbNewLine & "Account - " & newAccountName & Words.vLine(2) & _
        "Is this correct?"
    
    saidYes = MsgBox(message, vbYesNo, "Please Confirm") = vbYes
    
    If (Not saidYes) Then 'said No
    
        Call hideAndReload
    
    End If
    
End Sub

'@Info (validates to make sure user data can actually be used)
Private Sub verifyDataOnSubmit()
    
    Dim ws As Worksheet
    'verify the cmbAccountToEdit has a real account
    'verify the txt boxes have valid data
    'verify the txt boxes don't have already used data (mostly the tab)
    If oldAccountName = vNil Then 'They actually entered something in the cmboBox
    
        errorMsg "You did not select an Account to edit"
        
    ElseIf oldTabName = vNil Then 'Will find match if oldAccountName is valid and doesn't find a tab
    
        errorMsg "The Account you want to edit does not exist. Please check for capitalization and extra space"
        
    ElseIf newTabName = vNil Then
    
        errorMsg "You did not enter a new Abbreviation"
    
    ElseIf Len(newTabName) > 30 Then 'max tab name length
    
        errorMsg "Max length for the Abbreviation is 30 characters. Why would you put more???"
    
    ElseIf newAccountName = vNil Then
    
        errorMsg "You did not enter a new Account Name"
        
    ElseIf Left(newTabName, 1) = " " Then
    
        weirdDataMsg "New Abbreviation has a leading Space("" ""). Would you like to keep it?"
        
    ElseIf Right(newTabName, 1) = " " Then
    
        weirdDataMsg "New Abbreviation has a trailing Space("" ""). Would you like to keep it?"
    
    ElseIf Left(newAccountName, 1) = " " Then
    
        weirdDataMsg "New Account Name has a leading Space("" ""). Would you like to keep it?"
        
    ElseIf Right(newAccountName, 1) = " " Then
    
        weirdDataMsg "New Account Name has a leading Space("" ""). Would you like to keep it?"
    
    End If
    
    For Each ws In ThisWorkbook.Worksheets
        
        If ws.name = newTabName Then
            
            errorMsg "Abbreviation has already been taken"
            
        ElseIf ws.Range("A1").value = newAccountName Then
        
            errorMsg "Account Name is already taken"
            
        End If
        
    Next ws
    

End Sub

'@Info (Displays error message and refreshes form with user data unchanged)
Private Sub errorMsg(message As String)

    MsgBox message, vbCritical, "Invalid Data!"
    Call hideAndReload
    
End Sub
'@Info (Displays warning message about weird data like spaces)
Private Sub weirdDataMsg(message As String)

    Dim saidYes As Boolean
    
    saidYes = MsgBox(message, vbYesNo, "Are You Sure???")
    
    If Not saidYes Then
        Call hideAndReload
    End If
    
End Sub

'@Info (Refreshes form with user data unchanged)
Private Sub hideAndReload()

    frmEditAccountName.Hide
    frmEditAccountName.Show
    
End Sub

'@Info (Click Cancel - Closes form and shows message box informing of cancelled operation)
Private Sub btnCancel_Click()
    
    Unload Me
    
    MsgBox "No accounts edited", vbInformation, "No Changes"
    
    End
    
End Sub

'@Info (Form Shows - Tries to center it [not really working], calls 'Account To Edit' combobox populator)
Private Sub UserForm_Initialize()
    Dim lTop As Long, lLeft As Long
    Dim lRow As Long, lCol As Long
     
    With ActiveWindow.VisibleRange
        lRow = .Rows.count / 2
        lCol = .Columns.count / 2
    End With
     
    With Cells(lRow, lCol)
        lTop = .Top
        lLeft = .Left
    End With
     
    With Me
        .Top = lTop
        .Left = lLeft
    End With
    
    Call populateComboBox
    
End Sub

'@Info (Retrieves all account and tab names of valid sheets [first cell's comment Text = "S" or "M" for Semimonthly or Monthly, respectively] and stores them in cmbAccountToEdit)
Private Sub populateComboBox()

    Dim wb As Workbook, ws As Worksheet, tot As Worksheet
    Dim editType As String, wsTabName As String, wsAccountName As String
    
    Set wb = ThisWorkbook: Set tot = Config.getSheet_Total()
    
        For Each ws In wb.Worksheets
            
            wsTabName = ws.name
            wsAccountName = ws.Range("A1").value
            editType = editTypes.getEditType(wsTabName)
            
            If editTypes.isEditType(editType) Then
                
                Me.cmbAccountToEdit.AddItem wsAccountName
                
            End If
            
        Next ws
    
End Sub

'@Info (Click 'X' - close form, cancel operation)
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        End
    End If
End Sub
Private Sub cmbAccountToEdit_KeyDown(ByVal KeyCode As msforms.ReturnInteger, ByVal SHIFT As Integer)

If KeyCode = keys.enterKey() Then Call btnSubmit_Click

End Sub
Private Sub txtNewName_KeyDown(ByVal KeyCode As msforms.ReturnInteger, ByVal SHIFT As Integer)

If KeyCode = keys.enterKey() Then Call btnSubmit_Click

End Sub
Private Sub txtNewAbbrv_KeyDown(ByVal KeyCode As msforms.ReturnInteger, ByVal SHIFT As Integer)

If KeyCode = keys.enterKey() Then Call btnSubmit_Click

End Sub
