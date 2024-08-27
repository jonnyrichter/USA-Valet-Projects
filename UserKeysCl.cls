VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserKeysCl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub sendTab()
    Application.SendKeys "{tab}", True
End Sub

Public Sub sendText(text As String)

    Dim textLength As Long, i As Long
    
    textLength = Len(text)
    
    For i = 1 To textLength
        Application.SendKeys Mid(text, i, 1)
        Sleep 100
    Next i
    
End Sub
