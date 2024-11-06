import win32com.client
import traceback
import os

def print_excel(target_file_path, output_path):
    # 既にPDFが存在する場合は処理に失敗するので削除をしておく
    if os.path.exists(output_path):
        print(f"{target_file_path} -> deleting old pdf...")
        os.remove(output_path)

    excel = win32com.client.Dispatch("Excel.Application")
    # エクセルを展開
    print(f"{target_file_path} -> open file...")
    file = excel.Workbooks.Open(target_file_path)
    # 対象シートを指定
    file.WorkSheets(1).Select()
    print(f"{target_file_path} -> exporting...")
    # PDF出力 
    file.ActiveSheet.ExportAsFixedFormat(0, output_path)
    print(f"{target_file_path} -> completed!")

try:
    # 第1引数に開く既存エクセルの絶対パス、第2引数にPDFをエクスポートする絶対パスを指定する
    print_excel("/Users/tenhou/Desktop/test.xlsx", "/Users/tenhou/Desktop/output.pdf")
    print("completed!")
except:
    print('****** Exception ******')
    traceback.print_exc()