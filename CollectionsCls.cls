VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionsCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "CollectionCls"

Public Function toArray(col As Collection) As String()
    Dim output() As String
    Dim i As Integer
    
    If col.count > 0 Then
        ReDim output(col.count - 1) As String
        For i = 0 To col.count - 1
            output(i) = Replace(col.Item(i + 1), vbNewLine, vbNullString)
        Next i
        toArray = output()
    Else
        toArray = Arrays.strArray("") '//TODO: NO. THIS SHOULD BE AN ERROR
    End If
End Function
