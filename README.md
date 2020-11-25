# ツール名称
金融機関の明細収集ツール

# 概要
銀行口座、クレジットカード会社の明細データ（csvファイル）をエクセルファイルに取り込むpythonスクリプトです。

取得後、月末の明細行を判別して月毎の残高表を作成します。

水道光熱費の明細を判別して、事業と家事の按分計算用の表を作成します。（主に個人事業主向け用途）

## 画面イメージ

【口座残高画面】
![money-management-tool_1](https://user-images.githubusercontent.com/51041421/100288195-83b40380-2fb9-11eb-99d0-e00b109c2984.png)

【按分計算用画面】
![money-management-tool_2](https://user-images.githubusercontent.com/51041421/100288246-a0503b80-2fb9-11eb-8467-65a078dc4968.png)

【明細画面】
![money-management-tool_3](https://user-images.githubusercontent.com/51041421/100288265-aba36700-2fb9-11eb-9f27-185f6a76c66b.png)


# 動作環境

- Windows10 (Home ver.1909)
- Excel (Microsoft 365 ver.2020)
- Python3.7


# 機能詳細

## 対応している金融機関

### 銀行
- UFJ銀行
- 楽天銀行
- NEOBANK（旧SBI銀行）
- 三井住友銀行

### クレジットカード会社
- 楽天カード
- Viewカード
- SAISONカード


## ファイル・ディレクトリ構成

```
root:.
|   consts.py
|   export_for_accounting_firm.py
|   import_csv.py
|   main.bat
|   main.py
|   mmtUtil.py
|   reset_export_file.bat
|   reset_export_file.py
|   update_sheets.py
|
+---backup
|       Money-management_backup.xlsx
|
+---data
|       【Readme】サンプルデータファイルについて.txt
|       ※その他、各金融機関のサンプルデータ
|
+---output
       Money-management.xlsx
       支出管理_2020.xlsx
```

# 使い方

## main.bat

main.pyを実行するバッチファイルです。以下の内容を実行します。
1. バックファイルの作成（1世代のみ）  
作成先：「backup」フォルダ
2. csvデータの読み込みとSummaryシート等の更新  
取得元：「data」フォルダ
3. Excelファイルの保存  
保存先：「output」フォルダ
4. 会計事務所向けファイルの出力  
保存先：「output」フォルダ


## reset_export_file.bat
 reset_export_file.pyを実行するバッチファイルです。以下の内容を実行します。
1. Summaryシートの初期化
2. Divideシートの初期化
3. 各明細シートの初期化
4. Excelファイルの保存


# 補足
このソースは、pythonを使ってExcel業務の効率化を学ぶために作成したサンプルです。

## このサンプルで学べること

- csvファイルの読み込み
- Excelファイルへのデータ書き込み  
（日付、金額データを適切にデータ変換して取り込む）
- csvファイルの文字コードを判別  
(UTF-8とSJISを判別し、適切にファイルを開く)
- 和暦から西暦への変換
