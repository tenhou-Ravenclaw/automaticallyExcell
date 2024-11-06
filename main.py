from reportlab.lib.pagesizes import letter  # reportlabからページサイズをインポート
from reportlab.pdfgen import canvas  # reportlabからcanvasをインポート
from openpyxl import load_workbook  # Excelファイルを読み込むためにload_workbookをインポート

# 既存のExcelファイルのパス
file_path = '/Users/tenhou/Desktop/test.xlsx'  # Mac用のパスに変更

# Excelファイルを開く
workbook = load_workbook(file_path)  # 既存のExcelファイルを読み込む
worksheet = workbook.active  # 最初のワークシートを取得

