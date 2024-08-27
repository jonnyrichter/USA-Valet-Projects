VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPword 
   Caption         =   "Please Enter Passcode to Continue"
   ClientHeight    =   1425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4005
   OleObjectBlob   =   "frmPword.frx":0000
End
Attribute VB_Name = "frmPword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
End Sub

Private Sub btnSubmit_Click()
Dim x As Boolean
x = False

tryAgain:
If x = True Then
    frmPword.Hide
    frmPword.Show
End If
If Not Passwords.matchesPassword(Me.txtPword.value) Then
'If Me.txtPword <> "1022" And Me.txtPword <> "0324" Then
    MsgBox "Please try again.", vbCritical, "Incorrect"
    Me.txtPword = ""
    x = True
    GoTo tryAgain
End If

Unload Me

End Sub

Private Sub txtPword_KeyDown(ByVal KeyCode As msforms.ReturnInteger, ByVal SHIFT As Integer)

If KeyCode = keys.enterKey() Then Call btnSubmit_Click

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        End
    End If
End Sub
