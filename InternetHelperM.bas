Attribute VB_Name = "InternetHelperM"
Option Explicit

Private Const getClass As String = "InternetHelperM"
Public IE As InternetExplorer, HTML As HTMLDocument, ele As IHTMLElement, url As String

Private Declare PtrSafe Function InternetGetConnectedStateEx Lib "wininet.dll" _
    (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) _
        As Long
 
'Testing for internet connection
Public Function IsInternetConnected() As Boolean
 
    IsInternetConnected = InternetGetConnectedStateEx(0, "", 254, 0)
    
    If IsInternetConnected = False Then
        MsgBox "Please check connection and try again.", , "Internet Connection Not Established"
        End
    End If
 End Function

Public Sub OpenIE(url As String, Optional ThtreMode As Boolean = False)
Call IsInternetConnected

If IE Is Nothing Then Set IE = New InternetExplorer
IE.Visible = True '@Make this part of a settings form
If Not VersionControl.TestStatus Then
    IE.TheaterMode = ThtreMode
Else
    IE.TheaterMode = False
End If
IE.navigate url
WaitForM.BrowserToLoad
Set HTML = IE.Document
Err.clear
End Sub

Public Sub CloseIE()

On Error Resume Next
log.setClass(getClass).setMethod ("CloseIE")
IE.Quit
Set IE = Nothing
If Err <> 0 Then
    'debug.print "Internet Explorer window was already closed."
    Err.clear
End If
If Words.contains(Err.Description, "object variable") And Words.contains(Err.Description, "not set", vbTextCompare) Then
    log.warn "Tried to close IE when no instance detected."
End If

End Sub
Public Sub endAll()
    Selenide.CloseIE
    System.Update True
    End
End Sub
