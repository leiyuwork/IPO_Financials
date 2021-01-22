import pandas as pd
import os
import warnings

warnings.filterwarnings("ignore")

path = r'C:\Users\Ray94\Downloads\確認済み_566社\確認済み_566社'
files = os.listdir(path)
row_string = str(input("取り出したい項目名:"))

for file in files:
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    IS_data = pd.DataFrame(pd.read_excel(path + '\\' + file, sheet_name="Income Statement"))  # 获得每一个sheet中的内容
    IS_data = IS_data.replace("Restated\\n", "", regex=True)
    IS_data = IS_data.replace("Reclassified\\n", "", regex=True)
    IS_data = IS_data.replace("LTM\\n", "", regex=True)
    IS_data = IS_data.replace("12 months\\n", "", regex=True)

    Company_name = IS_data.iloc[3].values.tolist()
    Company = pd.DataFrame(Company_name).dropna()
    Company.to_csv(path + '\\' + "Fiscal_" + row_string + ".csv", mode='a', index=False, header=None,
                   date_format='%Y%m%d')

    Fiscal_year = IS_data.iloc[13].values.tolist()
    Fiscal_year_final = [Fiscal_year[:0:-1]]
    Fiscal = pd.DataFrame(Fiscal_year_final)
    Fiscal.iloc[0] = row_string + '_' + Fiscal.iloc[0].astype(str)
    print(Fiscal)
    Fiscal.to_csv(path + '\\' + "Fiscal_" + row_string + ".csv", mode='a', index=False, header=None, date_format='%Y%m%d')

    Item_value = IS_data[IS_data.iloc[:, 0].str.contains("Gross Profit", na=False)].values.tolist()
    for Values in Item_value:
        Item_value_final = [Values[:0:-1]]
    Item = pd.DataFrame(Item_value_final)
    Item.to_csv(path + '\\' + "Fiscal_" + row_string + ".csv", mode='a', index=False, header=None, date_format='%Y%m%d')