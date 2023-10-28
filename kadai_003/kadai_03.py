#coding:cp932

import pandas as pd
from glob import glob
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# ファイルパスを取得
filepaths = glob("データ/*.xlsx")

# データを読み込み、集計
data = []
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet1")
    data.append(df)
combined_df = pd.concat(data, ignore_index=True,axis=0)
grouped_df = combined_df.groupby(["商品", "売上年"]).agg({"金額（千円）": sum}).\
reset_index()

# 売上集計表作成Excelの作成
sales_totay = "売上集計表.xlsx"
wb = Workbook()
ws = wb.active

# 集計データのExcelへの転記
for row in dataframe_to_rows(combined_df, index=False, header=True):
    ws.append(row)

# ヘッダー部分のセルにグレーの塗りつぶしを適用
header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
for cell in ws["1:1"]:
    cell.fill = header_fill

# Excelファイルを保存
wb.save(sales_totay)

# ファイルを閉じる
wb.close()
