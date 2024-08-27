VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelenideCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const getClass As String = "SelenideCls"

Private Declare PtrSafe Function InternetGetConnectedStateEx Lib "wininet.dll" _
    (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) _
        As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" Alias "sleep" (ByVal dwMilliseconds As Long)
 
'Testing for internet connection
Public Function IsInternetConnected() As Boolean
 
    IsInternetConnected = InternetGetConnectedStateEx(0, vNil, 254, 0)
    
    If IsInternetConnected = False Then
        MsgBox "Please check connection and try again.", , "Internet Connection Not Established"
        System.Update True
        End
    End If
 End Function

Public Sub OpenIE(url As String, Optional seconds As Double = 1)
    Dim IE As InternetExplorer
    Call IsInternetConnected
    System.Update False
    Set IE = New InternetExplorer
    IE.Visible = True
    IE.navigate url
    WaitFor.setBrowser IE
    By.setBrowser IE
End Sub
Public Sub CloseIE()
    Dim IE As InternetExplorer
    Dim Shell As Object
    Set Shell = CreateObject("Shell.Application")
    For Each IE In Shell.Windows
        If TypeName(IE.Document) = "HTMLDocument" Then
            IE.Quit
            Set IE = Nothing
        End If
    Next
End Sub
Public Sub browserSleep(sleepTimeInMilliseconds As Long)
    Sleep sleepTimeInMilliseconds
End Sub



