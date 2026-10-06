# Excel 操作を Python でやってみよう

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Tomoaki-Kotsuka/python-kouza/blob/main/python_excel_kouza.ipynb)

Excel をふだん使っている人向けの Python 入門講座です。売上データの集計から、地区ごとのファイル分けまでを Python でやってみます。

## 始め方

1. 上の「Open in Colab」を押す(Google アカウントでログインしておく)
2. Colab のメニューで「ファイル → ドライブにコピーを保存」を押す
3. 上のセルから順に `Shift + Enter` で実行する

Excel データはノートブックが自動で用意するので、インストールやアップロードは要りません。

## ファイル

- `python_excel_kouza.ipynb` … 講座のノートブック
- `sales_data.xlsx` … 講座で使う売上データ(架空のデータです)
- `make_sales_data.py` … 売上データを作り直すスクリプト
- `split_by_area.txt` / `split_by_area.bas` / `PythonKouza.applescript` … 比較用の VBA
