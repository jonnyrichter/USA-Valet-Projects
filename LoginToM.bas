Attribute VB_Name = "LoginToM"
Option Explicit
'PercentByLogin is the first percentage that is used after routine has ended
Public Sub ADP()

    Const sleepTime As Long = 1000
    Const url = "https://runpayroll.adp.com/"
    Const userName_input_id As String = "login-form_username"
    Const next_button_id As String = "verifUseridBtn"
    Const password_input_id As String = "login-form_password"
    Const signIn_button_id As String = "signBtn"
    
    Dim userName_input As IHTMLElement
    Dim next_button As IHTMLElement
    Dim password_input As IHTMLElement
    Dim signIn_button As IHTMLElement

    Const uName = "usavaletparking" 'getUsername(Site)
    Dim pWord As String:  pWord = Ranges.getPasswordRange().value 'getPassword(Site)
    
    'Open login page
    InternetHelperM.OpenIE url, Not TestStatus
    
    'Activate the window by the title (don't think this is necessary but whatever)
    AppActivate HTML.Title
    
    WaitForM.ObjectById userName_input_id, 45
    
    Set userName_input = HTML.getElementById(userName_input_id)
    
    Sleep sleepTime
    
    userName_input.Click
    UserKeys.sendText uName
    
    Sleep sleepTime
    
    Set next_button = HTML.getElementById(next_button_id)
    next_button.Click
    
    Sleep sleepTime
    
    WaitForM.ObjectById password_input_id

    Set password_input = HTML.getElementById(password_input_id)
    
    Sleep sleepTime
    
    password_input.Click
    UserKeys.sendText pWord
    
    Sleep sleepTime
    
    Set signIn_button = HTML.getElementById(signIn_button_id)

    signIn_button.Click
    
    Sleep sleepTime
    
    WaitForM.ObjectByVisibleText "span", "Run Payroll"

End Sub

Public Sub ShiftPlanning(Optional Extension As String = vNil)

Dim ele_button As IHTMLElement

Const uName = "steven@usavaletparking.com" 'getUsername(Site)
Dim pWord As String: pWord = Ranges.getHumanityPasswordRange().value

InternetHelperM.OpenIE "https://usavalet.humanity.com" & Extension

If Not HTML.getElementById("userm") Is Nothing Then
    For Each ele In HTML.getElementsByTagName("script")
        If Words.contains(ele.innerText, uName, vbTextCompare) Then GoTo Success
    Next ele
    For Each ele In HTML.getElementsByTagName("a")
        If Words.contains(ele.getAttribute("onclick"), "dologout(0)") Then ele.Click: Exit For
    Next ele
    WaitForM.ObjectById "email"
End If

HTML.getElementById("email").value = uName
HTML.getElementById("Password").value = pWord
For Each ele_button In HTML.getElementsByName("login")
    If ele_button.innerText = "Log in" Then
        ele_button.Click
    End If
Next ele_button
Sleep 5000
InternetHelperM.OpenIE "https://usavalet.humanity.com" & Extension

Success:

End Sub
Public Function getShiftPlanningBaseUrl() As String
    getShiftPlanningBaseUrl = "https://usavalet.humanity.ShiftPlanning.com"
End Function


Public Sub SMSValet() 'This isn't currently needed

Const uName = "jon@usavaletparking.com" 'getUsername(Site)
Const pWord = "jon" 'getPassword(Site)

InternetHelperM.OpenIE "https://portal.smsvalet.com/"

HTML.getElementById("ctl00_ctl00_CphBodyCommon_CphBodyFront_TextBox1").value = uName
HTML.getElementById("ctl00_ctl00_CphBodyCommon_CphBodyFront_TextBox2").value = pWord
HTML.getElementById("ctl00_ctl00_CphBodyCommon_CphBodyFront_LsiASPxButton2").Click

WaitForM.BrowserToLoad

End Sub

