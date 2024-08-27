Attribute VB_Name = "LinesOfCodeM"
Option Explicit

Public Sub LinesOfCode()
Attribute LinesOfCode.VB_ProcData.VB_Invoke_Func = "L\n14"

Dim basModule As Variant
Dim LinesOfCode As Long

For Each basModule In Application.VBE.ActiveVBProject.VBComponents
    LinesOfCode = LinesOfCode + basModule.codemodule.CountOfLines
Next basModule

MsgBox "There are " & LinesOfCode & " Lines of Code in this Project" & Words.vLine(2) & _
    """I don't judge an engineer by how many lines of code he can write, but by how few.""" & Words.vLine() & _
    "-V. George", vbInformation, "Jon's Hard work"

End Sub
