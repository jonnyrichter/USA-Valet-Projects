VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegexCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "RegexCls"

Private reg As New RegExp
Private matcherValue As String

Public Function pattern(strPattern As String) As RegexCls
    With reg
        .Global = True
        .MultiLine = True
        .ignoreCase = False
        .pattern = strPattern
    End With
    Set pattern = Me
End Function

Public Function matcher(strMatcher As String) As RegexCls
    matcherValue = strMatcher
    Set matcher = Me
End Function

Public Function find() As Boolean
    find = True 'TODO
End Function

Public Function group() As String()
    Dim matches As Object

    If reg.test(matcherValue) Then
        Set matches = reg.Execute(matcherValue)
    Else
        Exit Function
    End If
    
    group = objects.toArray(matches)
End Function


