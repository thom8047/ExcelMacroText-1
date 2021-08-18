Private Sub Workbook_Open()
    Dim strMacroName As String
    
    strMacroName = CreateObject("WScript.Shell").Environment("process").Item("MacroName")
    If strMacroName <> "" Then Run strMacroName
End Sub
