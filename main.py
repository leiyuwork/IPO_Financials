import pandas as pd
import os
import warnings
import re

warnings.filterwarnings("ignore")

path_in = r'C:\Users\Ray94\Downloads\確認済み_566社\in'
path_out = r'C:\Users\Ray94\Downloads\確認済み_566社\out'
files = os.listdir(path_in)
row_string = str(input("取り出したい項目名:"))
sheet_string = str(input("取り出したいシート名:"))

Fiscal_year_final = []

for file in files:
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    IS_data = pd.DataFrame(pd.read_excel(path_in + '\\' + file, sheet_name=sheet_string))  # sheet内容取得
    IS_data = IS_data.replace("Restated\\n", "", regex=True)  # 不要情報を削除
    IS_data = IS_data.replace("Reclassified\\n", "", regex=True)  # 不要情報を削除
    IS_data = IS_data.replace("LTM\\n", "", regex=True)  # 不要情報を削除
    IS_data = IS_data.replace("12 months\\n", "", regex=True)  # 不要情報を削除

    Fiscal_year = IS_data.iloc[13].values.tolist()  # Fiscal情報を取り出す
    Fiscal_year.append('stock code')  # 後ろに　'stock code'　追加
    Fiscal_year_final.append(Fiscal_year[:0:-1])  # 順序逆転
Header = [max(Fiscal_year_final, key=len)]  # 一番多い情報年をheaderにする
Fiscal = pd.DataFrame(Header)  # dataframeの形にする
Fiscal.iloc[0] = row_string + '_' + Fiscal.iloc[0].astype(str)  # 項目名を各headerにつく
Fiscal.to_csv(path_out + '\\' + "Fiscal_" + row_string + ".csv", mode='a', index=False, header=None,
              date_format='%Y%m%d')  # CSVファイルのヘッダーが出来上がる

for file in files:
    print("処理中のファイルは" + file)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    IS_data = pd.DataFrame(pd.read_excel(path_in + '\\' + file, sheet_name="Income Statement"))  # sheet内容取得
    IS_data = IS_data.replace("Restated\\n", "", regex=True)  # 不要情報を削除
    IS_data = IS_data.replace("Reclassified\\n", "", regex=True)  # 不要情報を削除
    IS_data = IS_data.replace("LTM\\n", "", regex=True)  # 不要情報を削除
    IS_data = IS_data.replace("12 months\\n", "", regex=True)  # 不要情報を削除

    Company_name = IS_data.iloc[3].values.tolist()
    Stock_code = re.findall(r'(300[0-9][0-9][0-9])', Company_name[0])

    Item_value = IS_data[IS_data.iloc[:, 0].str.contains(row_string, na=False)].values.tolist()
    for Values in Item_value:
        try:
            Values.append(Stock_code[0])  # sheetから銘柄コード取得
        except:
            Stock_code = re.findall(r'(300[0-9][0-9][0-9])', file)  # 300104のように退場済み会社はシート内銘柄コードないため、ファイル名から取得
            Values.append(Stock_code[0])
        Item_value_final = [Values[:0:-1]]
    Item = pd.DataFrame(Item_value_final)
    Item.to_csv(path_out + '\\' + "Fiscal_" + row_string + ".csv", mode='a', index=False, header=None,
                date_format='%Y%m%d')
