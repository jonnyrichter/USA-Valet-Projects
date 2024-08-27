VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectsCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "ObjectsCls"

Public Function toArray(Obj As Object) As String()
    Dim output() As String
    Dim i As Integer
    
    If Obj.count > 0 Then
        ReDim output(Obj.count - 1) As String
        For i = 0 To Obj.count - 1
            output(i) = Replace(Obj.Item(i), vbNewLine, vbNullString)
        Next i
        toArray = output()
    Else
        toArray = Arrays.strArray("") '//TODO: NO. THIS SHOULD BE AN ERROR
    End If
End Function
