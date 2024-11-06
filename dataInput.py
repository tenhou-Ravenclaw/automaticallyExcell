from openpyxl import load_workbook 

def edit_excel_file():
    # Excelファイルを開く
    workbook = load_workbook('C:/Users/fromh/OneDrive/デスクトップ/test.xlsx')  # 既存のExcelファイルを読み込む
    worksheet = workbook.active  # 最初のワークシートを取得

    # セルの値を変更する例
    worksheet['A1'] = 'test'  # A1セルの値を変更
    worksheet['B1'] = 100  # B1セルの値を変更

    # 変更を保存
    workbook.save('C:/Users/fromh/OneDrive/デスクトップ/test.xlsx') 