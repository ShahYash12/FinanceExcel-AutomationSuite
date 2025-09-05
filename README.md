# 🧾 Finance Excel Automation with Python & VBA

This project automates the transformation of multiple `.xlsx` financial reports into macro-enabled `.xlsm` templates using Python and Excel VBA macros.

## 🔧 Project Features

- Batch process 100+ Excel files using Python (xlwings)
- Automatically copy all sheets or selected sheets into a template
- Retains formulas, formatting, images, and charts
- Built for a mid-sized **Accounting Firm** to simplify recurring finance reports
- Empowers non-technical users with **buttons** to trigger automation inside Excel

## 🧪 Files Included

- `macro_copy_selected_sheet.bas` – Copies selected sheet by prompt
- `macro_copy_all_sheets.bas` – Copies all sheets
- `macro_delete_extra_sheets.bas` – Deletes all but the first sheet
- `Template.xlsm` – Macro-enabled template (not included, upload your version)
- `Sample_Input.xlsx` – Sample input Excel (not included, upload your version)
- `your_script.py` – Python script to run the automation (see repo)

## ▶️ How to Run

1. Save your `.xlsx` files in a folder (e.g., `Input`)
2. Place `Template.xlsm` in a fixed folder
3. Define source and output folders when running the Python script
4. Use the buttons in Excel to:
   - Copy all sheets
   - Copy only selected sheets

## 🖱️ Macros Access

Make sure Excel **Trust Center Settings** are enabled:

- Enable **VBA Macros**
- Enable **Trust access to the VBA project object model**
- Mark file location as **Trusted**

## 💻 Requirements

- Python 3.x
- xlwings (`pip install xlwings`)
- Microsoft Excel (Windows only)

---

Built by [Yash Shah] for internal automation at a finance/accounting firm.
