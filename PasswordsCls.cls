VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PasswordsCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "PasswordsCls"

Private Const devPassword = "0324"
Private Const steveBauerPassword = "1022"

Public Function getDevPassword() As String
    getDevPassword = devPassword
End Function
Public Function matchesPassword(enteredPassword As String) As Boolean
    Dim pWords(1) As String 'add 1 to this array when a password is added
    Dim p As Variant
    Dim match As Boolean
    
    pWords(0) = steveBauerPassword
    pWords(1) = devPassword
    'pWords({^ + 1}) = "{newPWord}"
    For Each p In pWords
        If UCase(enteredPassword) = UCase(p) Then
            match = True
        End If
    Next p
    matchesPassword = match
End Function
