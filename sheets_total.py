# coding:utf-8
# 将工作表中多个工作表明细读取汇总

import pandas as pd
read_path = 'E:/data/11-内部计价/202012/111-内部计价登记表新-202012.xlsx'
save_path = 'E:/data/11-内部计价/2020内部计价登记表明细汇总.xlsx'


def sheet_detail(path, path_save):
    sheet_name = pd.ExcelFile(path).sheet_names
    df_details = pd.DataFrame([])
    for i in sheet_name:
        sub_detail = pd.read_excel(read_path, sheet_name=i)
        df_details = pd.concat([df_details, sub_detail])
    excel = pd.ExcelWriter(path_save)
    df_details.to_excel(excel)
    excel.save()


sheet_detail(read_path, save_path)
