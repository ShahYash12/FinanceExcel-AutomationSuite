
import xlwings as xw
import os

folder_path = "/Users/yashshah/Downloads/Shifting Project"  # <-- Change this

macro_code = '''
Sub ShiftContentRight_AllSheets()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim colCount As Integer
    Dim colCount1 As Integer
    colCount = 3  ' Number of columns to insert to shift content
    colCount1 = 1

    For Each ws In ThisWorkbook.Worksheets
        ' 1. Insert columns to shift cell content right
        ws.Columns("A").Resize(, colCount).Insert Shift:=xlToRight

        ' 2. Move each image right to match new columns
        For Each shp In ws.Shapes
            If shp.Type = msoPicture Then
                shp.Left = shp.Left + (colCount1 * 3# * ws.StandardWidth)    ' approx conversion
            End If
        Next shp
    Next ws
End Sub

Sub ShiftContentRight_SelectedSheet()
    Dim sheetName As String
    Dim ws As Worksheet
    Dim shp As Shape
    Dim colCount As Integer
    Dim colCount1 As Integer
    Dim sheetFound As Boolean

    colCount = 3  ' Number of columns to insert to shift content
    colCount1 = 1
    sheetFound = False

    ' Prompt user to enter sheet name
    sheetName = InputBox("Enter the name of the sheet you want to update:")

    ' Look for the sheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            sheetFound = True

            ' 1. Insert columns to shift cell content right
            ws.Columns("A").Resize(, colCount).Insert Shift:=xlToRight

            ' 2. Move each image right to match new columns
            For Each shp In ws.Shapes
                If shp.Type = msoPicture Then
                    shp.Left = shp.Left + (colCount1 * 3# * ws.StandardWidth)  ' approx conversion
                End If
            Next shp
            Exit For
        End If
    Next ws

    ' If sheet wasn't found
    If Not sheetFound Then
        MsgBox "Sheet '" & sheetName & "' not found!", vbExclamation
    End If
End Sub
'''

for file in os.listdir(folder_path):
    if file.endswith(".xlsx"):
        full_path = os.path.join(folder_path, file)
        new_file = full_path.replace(".xlsx", ".xlsm")
        wb.save(new_file)

        app = xw.App(visible=False)
        wb = app.books.open(full_path)

        # Save as .xlsm
        wb.api.SaveAs(new_file, FileFormat=52)
        wb.close()
        app.quit()

        print(f"✅ Converted: {file} → {os.path.basename(new_file)}")

# Inject macro + 2 buttons into all .xlsm files
for file in os.listdir(folder_path):
    if file.endswith(".xlsm"):
        full_path = os.path.join(folder_path, file)
        app = xw.App(visible=False)
        wb = app.books.open(full_path)

        try:
            wb.api.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString(macro_code)
        except Exception as e:
            print(f"⚠️ Could not inject macro into {file}: {e}")
            wb.close()
            app.quit()
            continue

        # Add two shape-buttons to first sheet
        sht = wb.sheets[0]
        try:
            btn1 = sht.api.Shapes.AddShape(1, 50, 50, 180, 30)  # Rectangle for All Sheets
            btn1.TextFrame2.TextRange.Text = "▶ Shift All Sheets"
            btn1.OnAction = "ShiftContentRight_AllSheets"

            btn2 = sht.api.Shapes.AddShape(1, 50, 100, 180, 30)  # Rectangle for Selected Sheet
            btn2.TextFrame2.TextRange.Text = "▶ Shift Selected Sheet"
            btn2.OnAction = "ShiftContentRight_SelectedSheet"
        except Exception as e:
            print(f"⚠️ Could not add shapes to {file}: {e}")

        wb.save()
        wb.close()
        app.quit()

        print(f"✅ Macro + Buttons added to: {file}")
