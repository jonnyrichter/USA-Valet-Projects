VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordsCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "WordsCls"

Public Function WordDif(String1 As Variant, String2 As Variant) As String
    
    Dim WordsA As Variant, WordsB As Variant
    Dim ndxA As Long, ndxB As Long, strTemp As String
    
    If String1 = String2 Then
        WordDif = vNil
        Exit Function
    End If
        
    WordsA = Split(String1, " ")
    WordsB = Split(String2, " ")
    
    For ndxB = LBound(WordsB) To UBound(WordsB)
        For ndxA = LBound(WordsA) To UBound(WordsA)
            If StrComp(WordsA(ndxA), WordsB(ndxB), vbTextCompare) = 0 Then
                WordsA(ndxA) = vbNullString
                Exit For
            End If
        Next ndxA
    Next ndxB
    
    For ndxA = LBound(WordsA) To UBound(WordsA)
        If WordsA(ndxA) <> vbNullString Then strTemp = strTemp & WordsA(ndxA) & " "
    Next ndxA
    
    WordDif = Trim(strTemp)

End Function

Public Function vTab(Optional numTabs As LongPtr = 1) As String
    vTab = vRep(vbTab, numTabs)
End Function

Public Function vLine(Optional numLines As LongPtr = 1) As String
    vLine = vRep(vbNewLine, numLines)
End Function

Public Function vRep(Repitition As String, Optional times As LongPtr = 1) As String
    Dim i As LongPtr, output As String

    For i = 1 To times
        output = output & Repitition
    Next i
    vRep = output
End Function

Public Function col(ColumnNumber As Long) As String
    Dim n As Long
    Dim c As Byte
    Dim s As String

    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    col = s
End Function

Public Function CharInst(expression As String, FindText As String) As Integer
    CharInst = (Len(expression) - (Len(Replace(expression, FindText, vNil)))) / Len(FindText)
End Function
Public Function contains(expression As String, find As String, Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    contains = IIf(InStr(1, expression, find, CompareMethod) > 0, True, False)
End Function
Public Function startsWith(expression As String, find As String, Optional ignoreCase As Boolean = False) As Boolean
    
    Dim eLength As Long, fLength As Long
    
    eLength = Strings.Len(expression): fLength = Strings.Len(find)
    If (fLength > eLength) Then
        startsWith = False
    Else
        If ignoreCase Then
            startsWith = Strings.Left$(Strings.UCase$(expression), fLength) = Strings.UCase$(find)
        Else
            startsWith = Strings.Left$(expression, fLength) = find
        End If
    End If
    
Exit Function

End Function
Public Function endsWith(expression As String, find As String, Optional ignoreCase As Boolean = False) As Boolean
    
    Dim eLength As Long, fLength As Long
    
    eLength = Strings.Len(expression): fLength = Strings.Len(find)
    If (fLength > eLength) Then
        endsWith = False
    Else
        If ignoreCase Then
            endsWith = Strings.Right$(Strings.UCase$(expression), fLength) = Strings.UCase$(find)
        Else
            endsWith = Strings.Right$(expression, fLength) = find
        End If
    End If
    
Exit Function
    
End Function
Public Function isEmpty(expression As String, ignoreWhiteSpace As Boolean) As Boolean
    If ignoreWhiteSpace Then ' " fsf" AND "asdf " AND " adf " AND "asdf" will evaluate the same
        isEmpty = Trim(expression) = vNil
    Else
        isEmpty = expression = vNil
    End If
End Function
Public Function formatStr(expression As String, ParamArray formatWith()) As String
    Const marker As String = "%s"
    Const newLine As String = "\n"
    Const newTab As String = "\t"
    Const newQuote As String = "\'"
    Dim i As Integer, f As Integer, e As Integer
    f = UBound(formatWith()) + 1
    e = Words.CharInst(expression, marker)
    If f <> e Then
        log.error Words.formatStr("NonEquivalentFormatException: Meant to replace %s words. Passed %s parameters.\nOffending Line:""%s""", e, f, expression)
        System.Update True
        End
    End If
    For i = 0 To UBound(formatWith())
        expression = Strings.Replace(expression, marker, formatWith(i), , 1)
    Next i
        expression = Strings.Replace(expression, newLine, vbNewLine)
        expression = Strings.Replace(expression, newTab, vbTab)
        expression = Strings.Replace(expression, "\'", """")
    formatStr = expression
End Function
