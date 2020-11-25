from chardet.universaldetector import UniversalDetector


# ファイルのエンコードタイプを取得
def get_encoding(file_path):
    detector = UniversalDetector()
    with open(file_path, mode='rb') as f:  # バイナリモードで開く
        for binary in f:
            detector.feed(binary)
            if detector.done:
                break
    detector.close()
    return detector.result['encoding']


# SMBCフォーマットの和暦から西暦へ変換
# SMBCの和暦フォーマット
# （例）平成28年10月4日・・・H28.10.04、令和2年1月10日・・・R02.01.10
def convert_sbmc_wareki_to_ad(wareki):
    reki_str = wareki[0]
    wareki_year = int(wareki[1:3])
    ad_year = 0
    month = wareki[4:6]
    day = wareki[7:9]

    if reki_str == "M":
        # 明治
        ad_year = 1867 + wareki_year
    elif reki_str == "T":
        # 大正
        ad_year = 1911 + wareki_year
    elif reki_str == "S":
        # 昭和
        ad_year = 1925 + wareki_year
    elif reki_str == "H":
        # 平成
        ad_year = 1988 + wareki_year
    elif reki_str == "R":
        # 令和
        ad_year = 2018 + wareki_year

    str_date = str(ad_year) + "/" + month + "/" + day
    return str_date
