import sys
import os
import shutil
import openpyxl
import consts
import import_csv
import update_sheets
import export_for_accounting_firm

# 実行
if __name__ == '__main__':

    # 出力用Excelファイルの存在確認
    if not os.path.exists(consts.EXCEL_FILE_NAME):
        print("File did not exists.")
        sys.exit()

    # Excelファイルのバックアップ（1世代のみ）
    shutil.copyfile(consts.EXCEL_FILE_NAME, consts.BACKUP_FILE_NAME)

    # Excelファイルを開く
    wb = openpyxl.load_workbook(consts.EXCEL_FILE_NAME)

    # インポートと出力
    import_csv.execute(wb)

    # 読み込みデータ解析・更新
    update_sheets.execute(wb)

    # Excelファイル保存
    wb.save(consts.EXCEL_FILE_NAME)

    # 会計事務所向けファイル出力
    export_for_accounting_firm.execute(wb)

    print("=== All done! ===")
