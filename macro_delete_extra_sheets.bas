Sub DeleteAllButFirstSheet()
    Dim i As Integer
    Application.DisplayAlerts = False
    For i = Sheets.Count To 2 Step -1
        Sheets(i).Delete
    Next i
    Application.DisplayAlerts = True
End Sub