import PySimpleGUI as sg
import os
import warnings
warnings.filterwarnings("ignore")
import pandas as pd
import re


sg.theme('DarkAmber')  # 设置当前主题
# 界面布局，将会按照列表顺序从上往下依次排列，二级列表中，从左往右依此排列
Sheets = ['Income Statement', 'Balance Sheet', 'Cash Flow', ]

layout = [[sg.Text('variable欄に改行を入れるとエラーになります。txt editorに保存してからの入力をお勧めします。')],
          [sg.Text('Input folder:'), sg.Input(), sg.FolderBrowse()],
          [sg.Text('Output folder:'), sg.Input(), sg.FolderBrowse()],
          [sg.Text('variable to generate:'), sg.Input()],
          [sg.Text('Click a sheet where the Variable is from')],
          [sg.Listbox(Sheets, key='-Sheet_list-', size=(20, 12))],
          [sg.Button('Ok'), sg.Button('Cancel')],
          [sg.Output(size=(60, 10))],
          [sg.Text('© 2021 Yu,lei.')], ]

# 创造窗口
window = sg.Window('S&P data generator(Yu Lei)', layout)
# 事件循环并获取输入值
while True:
    event, values = window.read()
    if event in (None, 'Cancel'):  # 如果用户关闭窗口或点击`Cancel`
        break
    print('You entered ', values[0], values[1], values[2], values["-Sheet_list-"][0])

    path_in = values[0]
    path_out = values[1]
    files = os.listdir(path_in)
    row_string = str(values[2])
    sheet_string = str(values["-Sheet_list-"][0])

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

        IS_data = pd.DataFrame(pd.read_excel(path_in + '\\' + file, sheet_name=sheet_string))  # sheet内容取得
        IS_data = IS_data.replace("Restated\\n", "", regex=True)  # 不要情報を削除
        IS_data = IS_data.replace("Reclassified\\n", "", regex=True)  # 不要情報を削除
        IS_data = IS_data.replace("LTM\\n", "", regex=True)  # 不要情報を削除
        IS_data = IS_data.replace("12 months\\n", "", regex=True)  # 不要情報を削除

        Company_name = IS_data.iloc[3].values.tolist()
        Stock_code = re.findall(r'(300[0-9][0-9][0-9])', Company_name[0])

        Item_value = IS_data[IS_data.iloc[:, 0].str.contains(row_string, na=False)].values.tolist()
        print(Item_value)
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
    sg.popup('処理完了')

window.close()
