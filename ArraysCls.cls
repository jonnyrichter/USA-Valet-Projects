VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArraysCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "ArraysCls"

Public Function isEmpty(ary() As String) As Boolean
isEmpty = Len(Join(ary())) = 0
End Function

Public Function removeByValue(StringArray() As String, values() As String) As String()

    Dim value As Variant
    Dim Str As Variant
    Dim col As New Collection
    Dim add As Boolean
    Dim s As Integer
    
    If isEmpty(StringArray()) Then Exit Function
    For Each Str In StringArray()
        add = True
        For Each value In values()
            If Str Like "*" & value & "*" Then
                add = False
                Exit For
            End If
        Next value
        If add Then
            col.add Str
        End If
    Next Str
    If col.count > 0 Then
        removeByValue = Collections.toArray(col)
    Else
        Exit Function
    End If

End Function

Public Function strArray(ParamArray values() As Variant) As String()
Dim v As Integer
Dim StringArray() As String
ReDim StringArray(UBound(values)) As String
For v = 0 To UBound(values())
    StringArray(v) = values(v)
Next v
strArray = StringArray
End Function

Public Function dblArray(ParamArray values() As Variant) As Double()
Dim v As Integer
Dim DoubleArray() As Double
ReDim DoubleArray(UBound(values)) As Double
For v = 0 To UBound(values())
    DoubleArray(v) = values(v)
Next v
dblArray = DoubleArray
End Function

Public Function lngArray(ParamArray values() As Variant) As LongPtr()
Dim v As Integer
Dim LongArray() As LongPtr
ReDim LongArray(UBound(values)) As LongLong
For v = 0 To UBound(values())
    LongArray(v) = values(v)
Next v
lngArray = LongArray
End Function

Public Function intArray(ParamArray values() As Variant) As Integer()
Dim v As Integer
Dim integerArray() As Integer
ReDim integerArray(UBound(values)) As Integer
For v = 0 To UBound(values())
    integerArray(v) = values(v)
Next v
intArray = integerArray
End Function

Public Sub printAll(arrayToPrint As Variant)

    Dim i As Long
    
    For i = 0 To UBound(arrayToPrint)
        Debug.Print "[" & i & "]: " & arrayToPrint(i)
    Next i
End Sub
