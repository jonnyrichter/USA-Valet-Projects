VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WaitForCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "WaitForCls"

Public found As Boolean
Public INC As Integer
Private HTML As HTMLDocument
Private IE As InternetExplorer

Public Sub setBrowser(browser As InternetExplorer)
    Set IE = browser
    BrowserToLoad
    Set HTML = IE.Document
End Sub
Private Function IELoaded()
    IELoaded = (IE.ReadyState = READYSTATE_COMPLETE And Not IE.Busy)
End Function

Public Sub BrowserToLoad(Optional seconds As Double = 1)
'Waits until the browser finishes after the intial ie.navigate - Only works on *initial* .navigate
    INC = 0
    Do
        Selenide.browserSleep seconds * 1000
        errorOut 15
    Loop While Not IELoaded()

End Sub
Private Sub errorOut(MaxIncrement As Integer)
    If INC = MaxIncrement Then
        log.error "Browser Errored out"
        Selenide.CloseIE
        MsgBox "Please try again later and/or notify a dev.", vbCritical, "There was a site-related error!"
        System.Update True
        End
    End If
End Sub


Public Sub ObjectsById(ParamArray ids())
'Waits until one of any objects (By ID) exists
    Dim ID As Variant
    INC = 0
    Do
        Selenide.browserSleep 500
        errorOut 30
        For Each ID In ids
            If Not HTML.getElementById(ID) Is Nothing And IELoaded() Then Exit Do
        Next ID
    Loop

End Sub

Public Sub ObjectByName(name As String, index As Integer)
'Waits until an object with the passed name and index exists on the page

    INC = 0
    Do
        Selenide.browserSleep 500
        errorOut 30
    Loop Until Not HTML.getElementsClickByName(name)(index) Is Nothing And IELoaded()

End Sub

Public Sub ObjectByPartialText(tagName As String, text As String)
'Waits until an object with the passed tagName exists with the given text
    Dim ele As IHTMLElement
    INC = 0: found = False
    Do
        Selenide.browserSleep 500
        For Each ele In HTML.getElementsByTagName(tagName)
            If Words.contains(ele.innerText, text) = True Then found = True: Exit For
        Next ele
        errorOut 30
    Loop Until found = True And IELoaded()

End Sub

Public Sub ObjectToDisappearById(ID As String)
    INC = 0
    On Error Resume Next
    Do
        Selenide.browserSleep 500
        errorOut 30
        If HTML.getElementById(ID) Is Nothing Then Exit Sub 'This may be the wrong way to go about it.
    Loop Until HTML.getElementById(ID).getAttribute("style").display = "none" And IELoaded()
End Sub

Public Sub ObjectInADPDetailFrame(ElementID As String, ElementTag As String) 'See if this can go by index
'Waits until an Iframe object exists by its ID
    Dim ElementCount As Integer
    Dim ele As IHTMLElement
    INC = 0
    Do
        Selenide.browserSleep 500
        errorOut 30
    Loop Until Not HTML.frames("iPDetail") Is Nothing And IELoaded()
    
    INC = 0
    found = False
    Do
        Selenide.browserSleep 500
        For Each ele In HTML.frames("iPDetail").Document.getElementsByTagName(ElementTag)
            If ele.ID = ElementID Then Exit Do
        Next ele
        errorOut 30
    Loop

End Sub

Public Sub ObjectInADPButtonFrame(ElementID As String, ElementTag As String) 'See if this can go by index
'Waits until an Iframe object exists by its ID
    Dim ElementCount As Integer
    Dim ele As IHTMLElement
    INC = 0
    Do
        Selenide.browserSleep 500
        errorOut 30
    Loop Until Not HTML.frames("iPButtons") Is Nothing And IELoaded()
    
    INC = 0
    Do
        Selenide.browserSleep 500
        For Each ele In HTML.frames("iPButtons").Document.getElementsByTagName(ElementTag)
            If ele.ID = ElementID Then Exit Do
        Next ele
        errorOut 30
    Loop

End Sub

Public Sub NextObjectByIdOrErrorMessage(NextObjectById As String, ParamArray ErrorObjectIds()) 'Might be specific to Google
    Dim ErrorObjectId As Variant 'For when error message doesn't exist until error
    INC = 0
    Do
        Selenide.browserSleep 1000
        errorOut 20
        For Each ErrorObjectId In ErrorObjectIds
            If Not HTML.getElementById(ErrorObjectId) Is Nothing Then
                If HTML.getElementById(ErrorObjectId).innerText <> vNil Then Exit Do
            End If
        Next ErrorObjectId
    Loop Until Not HTML.getElementById(NextObjectById) Is Nothing And IELoaded()
End Sub

Public Sub NextObjectByIdOrErrorMessageToDisplay(NextObjectById As String, ParamArray ErrorObjectIds()) 'Might be specific to Google
    Dim ErrorObjectId As Variant 'For when error message is display "none" until error
    INC = 0
    Do
        Selenide.browserSleep 1000
        errorOut 20
        For Each ErrorObjectId In ErrorObjectIds
            If Not HTML.getElementById(ErrorObjectId) Is Nothing Then 'Maybe add in by style attribute = display: none
                If HTML.getElementById(ErrorObjectId).getAttribute("style").display = "block" Then Exit Do
            End If
        Next ErrorObjectId
    Loop Until Not HTML.getElementById(NextObjectById) Is Nothing And IELoaded()
End Sub

Public Sub NextObjectByClassOrErrorMessage(NextObjectClassName As String, NextObjectClassIndex As Integer, ParamArray ErrorObjectIds()) 'Might be specific to Google
    Dim ErrorObjectId As Variant
    INC = 0
    Do
        Selenide.browserSleep 1000
        errorOut 20
        For Each ErrorObjectId In ErrorObjectIds
            If Not HTML.getElementById(ErrorObjectId) Is Nothing Then 'Maybe add in by style attribute = display: none
                If HTML.getElementById(ErrorObjectId).getAttribute("style") <> vNil Then
                    If HTML.getElementById(ErrorObjectId).getAttribute("style").display = "block" Then Exit Do
                Else
                    If HTML.getElementById(ErrorObjectId).innerText <> vNil Then Exit Do
                End If
            End If
        Next ErrorObjectId
    Loop Until Not HTML.getElementsByClassName(NextObjectClassName)(NextObjectClassIndex) Is Nothing And IELoaded()
End Sub

Public Sub ObjectToContainText(tagName As String, text As String)
    Dim ele As IHTMLElement
    INC = 0
    Do
        Selenide.browserSleep 1000
        errorOut 20
        For Each ele In HTML.getElementsByTagName(tagName)
            If Words.contains(ele.innerText, text, vbTextCompare) Then
                Exit Do
            End If
        Next ele
    Loop
End Sub
Public Sub ObjectByAttribute(tagName As String, attr As String, attributeValue As String)
    Dim ele As IHTMLElement
    INC = 0
    Do
        Selenide.browserSleep 1000
        errorOut 20
        For Each ele In HTML.getElementsByTagName(tagName)
            If Words.contains(Nz(ele.getAttribute(attr), vNil), attributeValue) Then
                Exit Sub
            End If
        Next ele
    Loop
End Sub
Public Sub elementEnabled(ele As IHTMLElement, Optional maxWait As Long = 20, Optional errorOut As Boolean = False)
    Dim waited As Long
    If ele.disabled = True Then 'These 3 if blocks make sure the object is displayed and enabled
        While ele.disabled = True
            Selenide.browserSleep 500
            waited = waited + 1
            If waited >= maxWait Then
                If errorOut Then
                    MsgBox "There was a site related Error!", vbCritical, "Error!"
                Else
                    Exit Sub
                End If
            End If
        Wend
    End If
    waited = 0
    If ele.getAttribute("type") = "hidden" Or ele.getAttribute("type") = "disabled" Then
        While ele.getAttribute("type") = "hidden"
            Selenide.browserSleep 500
            waited = waited + 1
            If waited >= maxWait Then
                If errorOut Then
                    MsgBox "There was a site related Error!", vbCritical, "Error!"
                Else
                    Exit Sub
                End If
            End If
        Wend
    End If
    waited = 0
    If Not ele.getAttribute("style") Is Nothing Then
        While ele.getAttribute("style").display <> "none" And ele.getAttribute("style").display <> vNil
            Selenide.browserSleep 500
            waited = waited + 1
            If waited >= maxWait Then
                If errorOut Then
                    MsgBox "There was a site related Error!", vbCritical, "Error!"
                Else
                    Exit Sub
                End If
            End If
        Wend
    End If
End Sub

