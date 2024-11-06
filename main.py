from reportlab.lib.pagesizes import letter  # reportlabからページサイズをインポート
from reportlab.pdfgen import canvas  # reportlabからcanvasをインポート
from openpyxl import load_workbook  # Excelファイルを読み込むためにload_workbookをインポート
import dataInput  # dataInput.pyをインポート
import print  # print.pyをインポート

# 既存のExcelファイルのパス
file_path = r'C:/Users/fromh/GitHub/automaticallyExcell/test.xlsx'  # Windows用のパスに変更

# Excelファイルを編集
dataInput.edit_excel_file(file_path)  # dataInput.pyのメソッドを呼び出す

# PDFを生成
pdf_path = r'C:\Users\tenhou\Desktop\output.pdf'  # PDFの保存先
canvas_obj = canvas.Canvas(pdf_path, pagesize=letter)  # PDFキャンバスを作成
canvas_obj.drawString(100, 750, "Excelファイルが編集されました。")  # テキストを追加
canvas_obj.save()  # PDFを保存

print("Excelファイルが編集され、PDFが生成されました。")
