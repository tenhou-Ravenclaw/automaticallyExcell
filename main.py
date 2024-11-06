import openpyxl
from xlsx2pdf import Xlsx2Pdf  # xlsx2pdfライブラリをインポート

# 既存のExcelファイルのパス
file_path = '/Users/tenhou/Desktop/test.xlsx'  # Mac用のパスに変更

# Excelファイルを開く
wb = openpyxl.load_workbook(file_path)
ws = wb.active

# セルに数字を代入
ws['A1'] = 123
ws['B1'] = 456

# 変更を保存
wb.save(file_path)

# PDFとして保存
pdf_path = '/Users/tenhou/Desktop/output.pdf'  # 保存先をDesktopに設定
xlsx2pdf = Xlsx2Pdf()  # Xlsx2Pdfのインスタンスを作成
xlsx2pdf.convert(file_path, pdf_path)  # ExcelファイルをPDFに変換
