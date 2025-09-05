Sub CopyAllSheets()
    Dim wb As Workbook, sht As Worksheet
    Set wb = Workbooks.Open(Application.GetOpenFilename())
    For Each sht In wb.Sheets
        sht.Copy Before:=ThisWorkbook.Sheets(1)
    Next sht
    wb.Close False
End Sub