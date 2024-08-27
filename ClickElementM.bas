Attribute VB_Name = "ClickElementM"
Option Explicit
Private i As Integer

Public Sub ByVisibleText(tagName As String, VisibleText As String)
i = 0
For Each ele In HTML.getElementsByTagName(tagName)
   If Words.contains(Trim(ele.innerText), VisibleText, vbTextCompare) Then Exit For
   i = i + 1
Next ele
HTML.getElementsByTagName(tagName)(i).Click
End Sub

Public Sub ByPartialOnClick(tagName As String, partialOnClick As String)
i = 0
For Each ele In HTML.getElementsByTagName(tagName)
    If Words.contains(ele.onClick, partialOnClick, vbTextCompare) Then Exit For
    i = i + 1
Next ele
HTML.getElementsByTagName(tagName)(i).Click
End Sub

Public Sub AllOfType(tagName As String)
For Each ele In HTML.getElementsByTagName(tagName)
    ele.Click
Next ele
End Sub

Public Sub byName(name As String, index As Integer)
HTML.getElementsByName(name)(index).Click
End Sub

Public Sub SelectOption(Dropdown As IHTMLElement, OptionByVisibleText As String)
For Each ele In Dropdown.Children
    If ele.innerText = OptionByVisibleText Then
        Dropdown.value = ele.value
        Exit For
    End If
Next ele
End Sub

Public Sub ByCssSelector(CssSelector As String) '@Test this
Dim Selectors() As String, tagName As String, Selector As Variant, a As Integer, Correct As Boolean
CssSelector = Replace(CssSelector, """]", "")
Selectors() = Split(CssSelector, "[")
tagName = Selectors(0)
Correct = True
i = 0
For Each ele In HTML.getElementsByTagName(tagName)
    For a = 1 To UBound(Selectors()) - 1
        If Mid(Selector(a), InStr(1, Selector(a), "=""") + 1) <> ele.getAttribute(Left(Selector(a), InStr(1, Selector(a), "=") - 1)) Then
        Correct = False: Exit For
        End If
    Next a
    If Correct = True Then Exit For
i = i + 1
Next ele
HTML.getElementsByTagName(tagName)(i).Click
End Sub

Public Sub ByAttribute(tagName As String, attr As String, attributeValue As String)

For Each ele In HTML.getElementsByTagName(tagName)
    If ele.getAttribute(attr) = attributeValue Then
        ele.Click
    End If
Next ele

End Sub



