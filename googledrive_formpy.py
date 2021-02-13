#googleコラボ用ソース


#インストール用
!pip install -U -q PyDrive
!pip install --upgrade -q gspread

#ライブラリ読み込み
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from google.colab import auth
from oauth2client.client import GoogleCredentials
import openpyxl
import pandas as pd
import gspread

#変数宣言（グローバル用）
auth.authenticate_user()
gauth = GoogleAuth()
gauth.credentials = GoogleCredentials.get_application_default()
drive = GoogleDrive(gauth)
gc = gspread.authorize(GoogleCredentials.get_application_default())
# ファイル名を指定してシートを開く
shtText2 = input('入力したいシート名を入力')#test_sheet
shtCategory = input('参照カテゴリは？タイムスタンプ:参加者氏名（所属と名前を記載してください）')
shtData = input('参照データは？')
sht = gc.open('Excel講習参加フォーム（回答） のコピー')
worksheet = sht.get_worksheet(0)
sht2 = gc.open(shtText2)
worksheet2 = sht2.get_worksheet(0)


# 読み込むセルの範囲の指定(読み込みたい範囲がわかっているのであれば適宜書き換える)
row_cnt = 27#worksheet.row_countここは行の数
col_cnt = 6#worksheet.col_countここは列の数

cells = worksheet.range(1, 1, row_cnt, col_cnt)

table_data = []
cols = []


#処理

for i, cell in enumerate(cells):
  cols.append(cell.value)
  if (i + 1) % col_cnt == 0:
    table_data.append(cols)
    cols = []

#table_dataここでデータの内容を確認できる。（一旦使わない）

#ここで取得したフォーム回答の内容についてのデータを配列でまとめる
df_list = []
for i,row in enumerate(table_data):
  if i == 0:
    header_cells = row
  else:
    row_dic = {}
    
    for k, v in zip(header_cells, row):
      row_dic[k] = v
    df_list.append(row_dic)
#print(df_list)

#ここで書き込みの処理を入れる。
setValue = ''
for j in df_list:
  if j[shtCategory] == shtData:
    print(j)
    setValue = j
    worksheet2.update_acell('A3', setValue['参加者氏名（所属と名前を記載してください）'])


