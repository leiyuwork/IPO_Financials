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

Header = [["Stock_code", "Fiscal_year", row_string]]
Set = pd.DataFrame(Header)
Set.to_csv(path_out + '\\' + "Fiscal_" + row_string + ".csv", mode='a', index=False, header=None,
           date_format='%Y%m%d')

for file in files:
    Final = []
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    IS_data = pd.DataFrame(pd.read_excel(path_in + '\\' + file, sheet_name=sheet_string))  # sheet内容取得
    """
    IS_data = IS_data.replace("Restated\\n", "", regex=True)  # 不要情報を削除
    IS_data = IS_data.replace("Reclassified\\n", "", regex=True)  # 不要情報を削除
    IS_data = IS_data.replace("LTM\\n", "", regex=True)  # 不要情報を削除

    """
    IS_data = IS_data.replace("12 months\\n", "", regex=True)  # 不要情報を削除
    Stock_code = re.findall(r'(300[0-9][0-9][0-9])', file)

    Fiscal_year_list = IS_data.iloc[13].values.tolist()  # Fiscal情報を取り出す
    del Fiscal_year_list[0]

    Item_value = IS_data[IS_data.iloc[:, 0].str.contains(row_string, na=False)].values.tolist()
    for Values in Item_value:
        del Values[0]

    for i in range(len(Fiscal_year_list)):
        Final.append([Stock_code[0], Fiscal_year_list[i], Values[i]])
    print(Final)

    Item = pd.DataFrame(Final)
    Item.to_csv(path_out + '\\' + "Fiscal_" + row_string + ".csv", mode='a', index=False, header=None,
                date_format='%Y%m%d')
