VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EditAccountNameDataStoreCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oldTabName As String
Private oldFullName As String
Private newTabName As String
Private newFullName As String

Public Sub setOldTabName(data As String)
    oldTabName = data
End Sub
Public Sub setOldFullName(data As String)
    oldFullName = data
End Sub
Public Sub setNewTabName(data As String)
    newTabName = data
End Sub
Public Sub setNewFullName(data As String)
    newFullName = data
End Sub
Public Function getOldTabName() As String
    getOldTabName = oldTabName
End Function
Public Function getOldFullName() As String
    getOldFullName = oldFullName
End Function
Public Function getNewTabName() As String
    getNewTabName = newTabName
End Function
Public Function getNewFullName() As String
    getNewFullName = newFullName
End Function
