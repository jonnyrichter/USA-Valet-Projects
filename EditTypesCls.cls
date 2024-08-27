VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EditTypesCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const edit As String = "E"
Private Const nonEdit As String = "N"

Public Function getEditType(sheetName As String) As String

    Dim commentText As String
    Dim sheetType As String
    
    commentText = Config.getSheet(sheetName).Range("A1").Comment.text
    
    sheetType = Strings.Split(commentText, ";")(1)
    
    getEditType = Strings.Replace(Strings.Trim$(sheetType), vbNewLine, vNil)

End Function

Public Function isEditSheet(sheetName As String) As Boolean
    isEditSheet = Me.isEditType(Me.getEditType(sheetName))
End Function
Public Function isEditType(sheetType As String) As Boolean
    isEditType = sheetType = edit
End Function

Public Function isNonEditSheet(sheetName As String) As Boolean
    isNonEditSheet = Me.isNonEditType(Me.getEditType(sheetName))
End Function
Public Function isNonEditType(sheetType As String) As Boolean
    isNonEditType = sheetType = nonEdit
End Function

