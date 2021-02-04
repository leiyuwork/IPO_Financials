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

window.close()
