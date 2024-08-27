Attribute VB_Name = "ChangeAuthentication"
Option Explicit

Sub ChangeADPpassword()

    Dim im As Worksheet
    Set im = Config.getSheet_Import()
    Dim OldPword As String
    Dim NewPword As String
    
    On Error GoTo endSub
    
    frmPword.Show
    OldPword = Ranges.getPasswordRange().value
    
    NewPword = InputBox("What has the new ADP Password been changed to?", "Password Change", OldPword)
    If NewPword = vbNullString Then
        MsgBox "Password not changed", , "Not Entered"
        Exit Sub
    End If
    
    System.unprotectSheet im
    Ranges.getPasswordRange().value = NewPword
    System.protectSheet im
    
    MsgBox "Password Changed", , "New Entry"
    
endSub:
    If Err <> 0 Then
        MsgBox "Error #: " & Err.Number & " - " & Err.Description, vbCritical
    End If

End Sub

Sub ChangeHumanityPassword()

    Dim im As Worksheet
    Set im = Config.getSheet_Import()
    Dim OldPword As String
    Dim NewPword As String
    
    On Error GoTo endSub
    
    frmPword.Show
    OldPword = Ranges.getHumanityPasswordRange()
    
    NewPword = InputBox("What has the Humanity Password changed to?", "Password Change", OldPword)
    If NewPword = vbNullString Then
        MsgBox "Password not changed", , "Not Entered"
        Exit Sub
    End If
    
    System.unprotectSheet im
    Ranges.getHumanityPasswordRange() = NewPword
    System.protectSheet im
    
    MsgBox "Password Changed", , "New Entry"
    
endSub:
    If Err <> 0 Then
        MsgBox "Error #: " & Err.Number & " - " & Err.Description, vbCritical
    End If

End Sub
