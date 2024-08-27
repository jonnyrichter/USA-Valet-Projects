VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "LoggerCls"

Private ofClass As String
Private ofMethod As String
Private count As Long

Public Function setClass(c As String) As LoggerCls
    ofClass = c
    count = 0
    Set setClass = Me
End Function
Public Function setMethod(m As String) As LoggerCls
    ofMethod = m
    count = 0
    Set setMethod = Me
End Function
Public Sub error(errorMsg As String)
    p "ERROR", errorMsg
End Sub
Public Sub trace(traceMsg As String)
    p "TRACE", traceMsg
End Sub
Public Sub warn(warnMsg As String)
    p "WARN", warnMsg
End Sub
Public Sub info(infoMsg As String)
    p "INFO", infoMsg
End Sub
Private Sub p(msgType As String, msg As String)
    count = count + 1
    If ofClass = vbNullString Then ofClass = "{Class not defined}"
    If ofMethod = vbNullString Then ofMethod = "{Method not defined}"
    Debug.Print Words.formatStr(count & ": %s [%s] %s.%s() -> %s", timeStamp(), msgType, ofClass, ofMethod, msg)
End Sub
Private Function timeStamp() As String
    timeStamp = Strings.Format$(Time(), "hh:mm:ss am/pm")
End Function

