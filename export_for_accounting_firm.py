import openpyxl
import consts


def execute(wb):
    sheets = wb.sheetnames

    # 不要シートを削除
    for sh_name in sheets:
        if sh_name not in consts.delete_list:
            wb.remove(wb[sh_name])

    # リネームして保存
    wb.save(consts.EXPORT_FILE_NAME)
