Attribute VB_Name = "WaitForM"
Option Explicit
Option Compare Text

Private Const Testing = False
Private Const DescError As String = "Browser Stalled. Please Try Again!"
Private Const TitleError As String = "Automation Error"
Public INC As Integer
Private found As Boolean

Public Sub BrowserToLoad()
'Waits until the browser finishes after the intial ie.navigate

INC = 0
Do
    Increment 1
    errorOut 30
Loop Until Not IE.Busy And IE.ReadyState = 4

End Sub

Public Sub ObjectById(ID As String, Optional maxWaitInSeconds As Integer = 20)
'Waits until an object with the passed ID exists on the page
INC = 0
While HTML.getElementById(ID) Is Nothing Or IE.ReadyState <> 4 Or IE.Busy
    Increment 1
    errorOut maxWaitInSeconds
Wend

End Sub
Public Sub EitherObjectById(Id1 As String, Id2 As String)
'Waits until one of two objects with the passed IDs exist on the page
INC = 0
Do
    Increment 1
    errorOut 20
Loop Until (Not HTML.getElementById(Id1) Is Nothing Or Not HTML.getElementById(Id2) Is Nothing) And IE.ReadyState = 4 And Not IE.Busy
End Sub

Public Sub ObjectByClass(className As String, Optional index As Integer = 0)
'Waits until an object with the passed className and index exists on the page
INC = 0
Do
    Increment 1
    errorOut 20
Loop Until Not HTML.getElementsByClassName(className)(index) Is Nothing And IE.ReadyState = READYSTATE_COMPLETE And Not IE.Busy
End Sub
Public Sub anyObjects(methods() As String, locators() As String)
    
    Dim objectElement As IHTMLElement
    Dim totalElements As Long
    Dim t As Long
    totalElements = UBound(methods())
    
    INC = 0
    t = -1
    Do
        t = t + 1
        If methods(t) = "id" Then
            Set objectElement = HTML.getElementById(locators(t))
        ElseIf methods(t) = "class" Then
            Set objectElement = HTML.getElementsByClassName(locators(t))(0)
        ElseIf methods(t) = "name" Then
            Set objectElement = HTML.getElementsByName(locators(t))(0)
        End If
        If t = totalElements - 1 Then t = -1
        Increment 1
        errorOut 20
    Loop Until Not objectElement Is Nothing
End Sub
Public Sub ObjectByName(name As String, Optional index As Integer = 0)
'Waits until an object with the passed name and index exists on the page
INC = 0
Do
    Increment 1
    errorOut 20
Loop Until Not HTML.getElementsByName(name)(index) Is Nothing And IE.ReadyState = 4 And Not IE.Busy

End Sub

Public Sub ObjectByVisibleText(tagName As String, text As String)
'Waits until an object with the passed tagName exists with the given text
INC = 0: found = False
Do
    Increment 1
    For Each ele In HTML.getElementsByTagName(tagName)
        If Words.contains(ele.innerText, text, vbTextCompare) Then found = True: Exit For
    Next ele
    errorOut 60
Loop Until found = True And IE.ReadyState = 4 And Not IE.Busy

End Sub

Public Sub HTMLObject(Element As IHTMLElement) 'This doesn't work
'If all else fails, wait for the actual object
INC = 0
Do
    Increment 1
    errorOut 20
Loop While Element Is Nothing Or IE.ReadyState <> 4 Or IE.Busy

End Sub

Public Sub ObjectToDisappearById(ID As String)
INC = 0
Do
    Increment 1
    errorOut 30
Loop Until HTML.getElementById(ID).getAttribute("style").display = "none" And IE.ReadyState = READYSTATE_COMPLETE And Not IE.Busy
End Sub

Public Sub ObjectToHaveText(objectElement As IHTMLElement, text As String)
INC = 0
Do
    Increment 1
    errorOut 20
Loop Until Words.contains(objectElement.innerText, text, vbTextCompare) And IE.ReadyState = READYSTATE_COMPLETE And Not IE.Busy
End Sub
Private Sub Increment(SecondsToWait As Integer)
If Testing = False Then
    Application.Wait Now() + TimeValue("00:00:" & SecondsToWait)
    DoEvents
    INC = INC + 1
End If
End Sub
Private Sub errorOut(MaxIncrement As Integer)
'log.trace INC
If INC = MaxIncrement Then
    System.Update True
    IE.Quit
    MsgBox DescError, vbCritical, TitleError
    End
End If
End Sub
