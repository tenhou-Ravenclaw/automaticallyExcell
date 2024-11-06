import openpyxl
import win32com.client

# 既存のExcelファイルのパス
existing_excel_path = 'existing_file.xlsx'

# Excelファイルを開く
wb = openpyxl.load_workbook(existing_excel_path)
ws = wb.active

# セルに数字を代入
ws['A1'] = 123
ws['B1'] = 456

# 変更を保存
wb.save(existing_excel_path)

# Excelアプリケーションを起動
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

# Excelファイルを開く
workbook = excel.Workbooks.Open(existing_excel_path)

# PDFとして保存
pdf_path = 'output.pdf'
workbook.ExportAsFixedFormat(0, pdf_path)

# Excelを閉じる
workbook.Close(SaveChanges=False)
excel.Quit()

print(f"PDFとして保存しました: {pdf_path}")
