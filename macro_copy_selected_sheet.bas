Sub CopySelectedSheet()
    Dim sheetName As String
    sheetName = InputBox("Enter the name of the sheet to copy:")
    If sheetName = "" Then Exit Sub
    On Error Resume Next
    Sheets(sheetName).Copy Before:=Sheets(1)
    If Err.Number <> 0 Then MsgBox "Sheet not found!"
End Sub