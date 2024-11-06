from openpyxl import load_workbook 

def edit_excel_file(file_path):
    # Excelファイルを開く
    workbook = load_workbook(file_path)  # 既存のExcelファイルを読み込む
    worksheet = workbook.active  # 最初のワークシートを取得

    # セルの値を変更する例
    worksheet['A1'] = '新しい値'  # A1セルの値を変更
    worksheet['B1'] = 100  # B1セルの値を変更

    # 変更を保存
    workbook.save(file_path)  # 同じファイルに保存

if __name__ == "__main__":
    file_path = r'file_path'  # Excelファイルのパス（Windows用）
    edit_excel_file(file_path)  # メソッドを呼び出す