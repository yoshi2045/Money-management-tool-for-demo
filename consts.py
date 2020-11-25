#
# 定数定義
#

# 読み込み対象データフォルダ
DATA_PATH = "./data"

# 出力先エクセルファイル名
EXCEL_FILE_NAME = "./output/Money-management.xlsx"

# バックアップファイル名
BACKUP_FILE_NAME = "./backup/Money-management_backup.xlsx"

# 会計事務所向けファイル名
EXPORT_FILE_NAME = "./output/支出管理.xlsx"

# シート名
SH_SUMMARY_BANK = "Summary_Bank"
SH_DIVIDE = "Divide"
SH_BANK_UFJ_FAMILY = "Bank_UFJ(Family)"
SH_BANK_UFJ_BIZ = "Bank_UFJ(Biz)"
SH_BANK_RAKUTEN = "Bank_Rakuten"
SH_BANK_NEO = "Bank_NEO"
SH_BANK_SMBC = "Bank_SMBC"
SH_CARD_RAKUTEN = "Card_Rakuten"
SH_CARD_VIEW = "Card_View"
SH_CARD_SAISON = "Card_SAISON"

# 出力対象シート名
delete_list = ["Divide",
               "Bank_UFJ(Biz)",
               "Bank_Rakuten",
               "Card_Rakuten"]
