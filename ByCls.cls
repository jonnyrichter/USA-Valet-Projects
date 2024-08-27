VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ByCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public objectNotFound As Boolean
Private Const getClass As String = "ByCls"

Private Const sleepTime As Integer = 500
Private htmlFrame As HTMLDocument
Private htmlFrameId As String
Private sqlQuery As String
Private resetFrame As Boolean
Private HTML As HTMLDocument
Private IE As InternetExplorer
Private linkToClickOnFailure As String

Public Sub openUrl(url As String)
    IE.navigate url
    WaitFor.BrowserToLoad
End Sub
Public Sub setBrowser(browser As InternetExplorer)
    Set IE = browser
    Set HTML = IE.Document
End Sub
Public Sub resetFrames()
    Set htmlFrame = Nothing
End Sub
Public Sub setFrame(frameDoc As HTMLDocument)
    Set htmlFrame = frameDoc
End Sub
Public Function frameIsNothing() As Boolean
    frameIsNothing = htmlFrame Is Nothing
End Function
Private Function ByBy(partialMatch As Boolean, identifierEquals As String, Optional tagNamesCommaDelim As String = "*", Optional index As Integer = 0, Optional waitUntil As Boolean = True, Optional maxWait As Integer = 15, Optional waitEnabled As Boolean = True) As IHTMLElement
    Dim n As Integer
    Dim ele As IHTMLElement
    Dim ElCol As IHTMLElementCollection 'IHTMLElementCollection
    Dim match As Boolean
    Dim identifier() As String
    Dim tagNames() As String
    Dim tagName As Variant
    Dim value As String
    Dim t As Integer
    
    identifier() = Strings.Split(identifierEquals, "=")
    tagNames() = Strings.Split(tagNamesCommaDelim, ",")
    
    Do
        For Each tagName In tagNames()
            'objectNotFound = True 'What am I using this for?
            If CStr(tagName) = "*" Then
                If htmlFrame Is Nothing Then
                    Set ElCol = HTML.all()
                Else
                    Set ElCol = htmlFrame.all()
                End If
            Else
                If htmlFrame Is Nothing Then
                    Set ElCol = HTML.getElementsByTagName(CStr(tagName))
                Else
                    Set ElCol = htmlFrame.getElementsByTagName(CStr(tagName))
                End If
            End If
            
            For Each ele In ElCol
                If identifier(0) = "text()" Then
                    'On Error Resume Next
                    value = Nz(ele.textContent, vNil)
                Else 'attribute
                    On Error Resume Next
                    value = Nz(ele.getAttribute(identifier(0)), vNil)
                    If Err <> 0 And Not Words.contains(Err.Description, "Access is denied") Then
                        MsgBox Err.Number & ": " & Err.Description
                        GoTo timeOut
                    Else
                        Err.clear
                        On Error GoTo 0
                    End If
                End If
                If partialMatch Then
                    match = Words.contains(value, identifier(1))
                Else
                    match = value = identifier(1)
                End If
                
                If match And n = index Then
                    If waitEnabled Then WaitFor.elementEnabled ele
                    Set ByBy = ele
                    If resetFrame Then Set htmlFrame = Nothing
                    Exit Function
                ElseIf match Then
                    n = n + 1
                End If
            Next ele
        Next tagName
        
        Selenide.browserSleep (sleepTime) 'Make this use the timeout methods
        t = t + 1
        If t > maxWait And waitUntil Then 'TimeOutException
timeOut:
            log.error "Not found By.xpath: { " & IIf(index > 0, "(", vNil) & "//" & tagNamesCommaDelim & "[" & Replace(identifierEquals, "=", "='") & "']" & IIf(index > 0, "[" & index + 1 & "])", vNil) & " }"
            System.Update True
            Selenide.CloseIE
            MsgBox "Please try again later and/or notify a dev.", vbCritical, "There was a site-related error!"
            End
        ElseIf t > maxWait Then
            Set ByBy = Nothing
            Exit Function
        End If
    Loop While waitUntil 'Object not found yet; loop until it is
End Function

Public Function ID(byId As String, Optional waitUntil As Boolean = True, Optional maxWait As Integer = 15, Optional waitEnabled As Boolean = True) As IHTMLElement
    Dim t As Integer
    Do
        If htmlFrame Is Nothing Then
            Set ID = HTML.getElementById(byId)
        Else
            Set ID = htmlFrame.getElementById(byId)
        End If
        If Not ID Is Nothing Then
            If waitEnabled Then WaitFor.elementEnabled ID
            If resetFrame Then Set htmlFrame = Nothing
            Exit Function
        End If
        Selenide.browserSleep sleepTime
        t = t + 1
        If t = maxWait * 2 Then 'TimeOutException
            log.error "Not found By.id { " & byId & " }"
            If linkToClickOnFailure <> vNil Then
                By.linkText(linkToClickOnFailure).Click
                WaitFor.BrowserToLoad
            End If
            System.Update True
            Selenide.CloseIE
            MsgBox "Please try again later and/or notify a dev.", vbCritical, "There was a site-related error!"
            End
        End If
    Loop While waitUntil
End Function

Public Function name(byName As String, Optional index As Integer = 0, Optional waitUntil As Boolean = True, Optional waitEnabled As Boolean = True) As IHTMLElement
    Dim t As Integer
    Do
        If htmlFrame Is Nothing Then
            Set name = HTML.getElementsByName(byName)(index)
        Else
            Set name = htmlFrame.getElementsByName(byName)(index)
        End If
        If Not name Is Nothing Then
            If waitEnabled Then WaitFor.elementEnabled name
            If resetFrame Then Set htmlFrame = Nothing
            Exit Function
        End If
        Selenide.browserSleep sleepTime
        t = t + 1
        If t = 15 Then 'TimeOutException
            log.error "Not found By.name { " & byName & " } index { " & index & " }"
            System.Update True
            Selenide.CloseIE
            MsgBox "Please try again later and/or notify a dev.", vbCritical, "There was a site-related error!"
            End
        End If
    Loop While waitUntil
End Function
Public Function classCollection(className As String) As IHTMLElementCollection
    Set classCollection = HTML.getElementsByClassName(className)
End Function
Public Function tagCollection(tagName As String) As IHTMLElementCollection
    Set tagCollection = HTML.getElementsByTagName(tagName)
End Function
Public Function nameCollection(nameValue As String) As IHTMLElementCollection
    Set nameCollection = HTML.getElementsByName(nameValue)
End Function
Public Function className(byClassName As String, Optional index As Integer = 0, Optional waitUntil As Boolean = True, Optional waitEnabled As Boolean = True) As IHTMLElement
    Dim t As Integer
    Do
        If htmlFrame Is Nothing Then
            Set className = HTML.getElementsByClassName(byClassName)(index)
        Else
            Set className = htmlFrame.getElementsByClassName(byClassName)(index)
        End If
        If Not className Is Nothing Then
            If waitEnabled Then WaitFor.elementEnabled className
            If resetFrame Then Set htmlFrame = Nothing
            Exit Function
        End If
        Selenide.browserSleep sleepTime
        t = t + 1
        If t = 15 Then 'TimeOutException
            log.error "Not found By.className { " & byClassName & " } index { " & index & " }"
            System.Update True
            Selenide.CloseIE
            MsgBox "Please try again later and/or notify a dev.", vbCritical, "There was a site-related error!"
            End
        End If
    Loop While waitUntil
End Function

Public Function tagName(byTagName As String, Optional index As Integer = 0, Optional waitUntil As Boolean = True, Optional waitEnabled As Boolean = True) As IHTMLElement
    Dim t As Integer
    Do
        Set tagName = HTML.getElementsByTagName(byTagName)(index)
        If Not tagName Is Nothing Then
            If waitEnabled Then WaitFor.elementEnabled tagName
            If resetFrame Then Set htmlFrame = Nothing
            Exit Function
        End If
        Selenide.browserSleep sleepTime
        t = t + 1
        If t = 15 Then 'TimeOutException
            log.error "Not found By.id { " & byTagName & " } index { " & index & " }"
            System.Update True
            Selenide.CloseIE
            MsgBox "Please try again later and/or notify a dev.", vbCritical, "There was a site-related error!"
            End
        End If
    Loop While waitUntil
End Function

Private Function text(partialMatch As Boolean, byText As String, Optional tagNamesCommaDelim As String = "*", Optional index As Integer = 0, Optional waitUntil As Boolean = True) As IHTMLElement
    Set text = ByBy(partialMatch, "text()=" & byText, tagNamesCommaDelim, index, waitUntil)
End Function

Private Function byAttr(partialMatch As Boolean, attributeName As String, attributeValue As String, Optional tagNamesCommaDelim As String = "*", Optional index As Integer = 0, Optional waitUntil As Boolean = True) As IHTMLElement
    Set byAttr = ByBy(partialMatch, attributeName & "=" & attributeValue, tagNamesCommaDelim, index, waitUntil)
End Function

Public Function attr(attributeName As String, exactAttributeValue As String, Optional tagNamesCommaDelim As String = "*", Optional index As Integer = 0, Optional waitUntil As Boolean = True) As IHTMLElement
    Set attr = byAttr(False, attributeName, exactAttributeValue, tagNamesCommaDelim, index, waitUntil)
End Function

Public Function partialAttribute(attributeName As String, partialAttributeValue As String, Optional tagNamesCommaDelim As String = "*", Optional index As Integer = 0, Optional waitUntil As Boolean = True) As IHTMLElement
    Set partialAttribute = byAttr(True, attributeName, partialAttributeValue, tagNamesCommaDelim, index, waitUntil)
End Function

Public Function onClick(exactOnClickValue As String, Optional tagNamesCommaDelim As String = "*", Optional index As Integer = 0, Optional waitUntil As Boolean = True) As IHTMLElement
    Set onClick = byAttr(False, "onclick", exactOnClickValue, tagNamesCommaDelim, index, waitUntil)
End Function

Public Function partialOnClick(partialOnClickValue As String, Optional tagNamesCommaDelim As String = "*", Optional index As Integer = 0, Optional waitUntil As Boolean = True) As IHTMLElement
    Set partialOnClick = byAttr(True, "onclick", partialOnClickValue, tagNamesCommaDelim, index, waitUntil)
End Function

Public Function partialText(textToFind As String, Optional tagNamesCommaDelim As String = "*", Optional index As Integer = 0, Optional waitUntil As Boolean = True) As IHTMLElement
    Set partialText = text(True, textToFind, tagNamesCommaDelim, index, waitUntil)
End Function

Public Function exactText(textToFind As String, Optional tagNamesCommaDelim As String = "*", Optional index As Integer = 0, Optional waitUntil As Boolean = True) As IHTMLElement
    Set exactText = text(False, textToFind, tagNamesCommaDelim, index, waitUntil)
End Function

Public Function linkText(textToFind As String, Optional index As Integer = 0, Optional waitUntil As Boolean = True) As IHTMLElement
    Set linkText = exactText(textToFind, "a", index, waitUntil)
End Function

Public Function partialLinkText(textToFind As String, Optional index As Integer = 0, Optional waitUntil As Boolean = True) As IHTMLElement
    Set partialLinkText = partialText(textToFind, "a", index, waitUntil)
End Function

Public Function buttonText(textToFind As String, Optional index As Integer = 0, Optional waitUntil As Boolean = True) As IHTMLElement
    Set buttonText = exactText(textToFind, "button", index, waitUntil)
End Function

Public Function partialButtonText(textToFind As String, Optional index As Integer = 0, Optional waitUntil As Boolean = True) As IHTMLElement
    Set partialButtonText = partialText(textToFind, "button", index, waitUntil)
End Function

Public Function frameById(frameId As String, Optional waitUntil As Boolean = True, Optional allowFrameReset As Boolean = True) As ByCls
    Dim t As Integer
    Dim f As Integer
    Dim fra As IHTMLElement
    Dim frameCollection As IHTMLElementCollection
    
    resetFrame = allowFrameReset
    
    WaitFor.BrowserToLoad
    Do
    DoEvents
tryAgain:
        f = 0
        Set frameCollection = HTML.getElementsByTagName("iframe")
        For Each fra In frameCollection 'get the iframe collection
            If fra.ID = frameId Then 'see if the object's id matches the frame we're looking for
                Set htmlFrame = fra
                Set frameById = Me
                Exit Function
            End If
            f = f + 1
        Next fra
        If f > frameCollection.length - 1 Then
            Selenide.browserSleep sleepTime
            t = t + 1
            If t = 15 Then 'TimeOutException
                log.error "Not found By.frameId { " & frameId & " }"
                System.Update True
                Selenide.CloseIE
                'Set IE = Nothing
                End
            End If
            GoTo tryAgain
        End If
        
    Loop While waitUntil
End Function

Public Sub SelectOption(Dropdown As IHTMLElement, optionText As String)
    Dim ele As IHTMLElement
    For Each ele In Dropdown.Children
        If ele.innerText = optionText Then
            Dropdown.value = ele.getAttribute("value")
            Exit For
        End If
    Next ele
End Sub

Public Sub SelectOptionByPartial(Dropdown As IHTMLElement, OptionClickPartialText As String)
    Dim ele As IHTMLElement
    For Each ele In Dropdown.Children
        If Words.contains(ele.innerText, OptionClickPartialText) Then
            Dropdown.value = ele.getAttribute("value")
            Exit For
        End If
    Next ele
End Sub

Public Sub focusSetInnerText(tagName As String, attri As String, attributeValue As String, innerTextToSet As String) 'Maybe Google Cal specific
    Dim i As Integer
    Dim ele As IHTMLElement
    Set ele = By.attr(attri, attributeValue, tagName)
    ele.focus
    ele.innerText = innerTextToSet
End Sub

Public Sub setInnerTextByPartialAttribute(tagName As String, attri As String, attributeValue As String, innerTextToSet As String) 'Maybe Google Cal specific
    Dim i As Integer
    Dim ele As IHTMLElement
    Set ele = By.partialAttribute(attri, attributeValue, tagName)
    ele.focus
    ele.innerText = innerTextToSet
End Sub
Public Function IsUsernameDisplayed(Username As String) As Boolean '@make sure this works
    Dim ele As IHTMLElement
    For Each ele In HTML.all()
        If Words.contains(ele.innerText, Username, vbTextCompare) Then
            IsUsernameDisplayed = True
            Exit Function
        End If
    Next ele
End Function
Public Sub clickLinkOnFailure(link As String)
    linkToClickOnFailure = link
End Sub
