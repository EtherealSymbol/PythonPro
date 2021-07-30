import numpy as np
import pandas as pd


def extract_20196():
    file_name = 'Cost_in2019/2019年6月产品成本计算单查询.xls'
    # file_name = '2019年6月产品成本计算单查询.xls'
    df = pd.read_excel(file_name, sheet_name="Sheet")
    # print(df)

    # print(df.values)
    # for row in df.values:
    #     print(row)

    # 将所有数据提取出来存到字典中
    dict_data = {}

    '''
    数据结构：
    dict_data:{
        '101000105' :{
            '201900008': {
                '物料': '二羟丙茶碱',
                '物料规格': '千克',
                ...
            },
            ...
        },
        ...
        dict_sno: {
            list_row: {},
            list_row: {},
            ...
        },
        dict_sno: {
            list_row: {},
            list_row: {},
            ...
        },
        ...

    }
    '''

    # for row in df['产品编号']:
    #     # 布尔取反
    #     if 1 - np.isnan(row):
    #         # row是101000338.0这类的数字
    #         # 转化成字符串并去掉后面的.0
    #         str_cp_no = str(row)[0:-1]
    #         # print(str_cp_no)
    #         # 出现新的编号
    #         if dict_data.get(str_cp_no) is not None:
    #             dict_son = {}
    #             list_row = []
    df_values = df.values
    len_values = len(df_values)
    for i in range(len_values):
        # print(df_values[i])
        # print(df['产品编号'][i], df['物料'][i])
        # 布尔取反
        row_cp_bh = df['产品编号'][i]
        # 费用项目这一行，确定是原材料才行
        row_fy_xm = df['费用项目'][i]
        if 1 - np.isnan(row_cp_bh) and row_fy_xm == "原材料":
            # print(df_values[i])
            # row是101000338.0这类的数字
            # 转化成字符串并去掉后面的.0
            # 产品编号
            str_cp_no = str(row_cp_bh)[0:-2]
            # 对应的物料编号
            row_wl_bh = df['物料编号'][i]
            str_wl_no = str(row_wl_bh)[0:-2]
            # print(str_cp_no, str_wl_no, df['物料'][i], df['物料规格'][i], df['物料计量单位'][i], df['投入数量'][i],
            #       df['投入金额'][i], df['数量'][i], df['单价'][i], df['金额'][i])
            list_row = {'物料': df['物料'][i],
                        '物料规格': df['物料规格'][i],
                        # '物料计量单位': df['物料计量单位'][i],
                        '投入数量': df['投入数量'][i],
                        '投入金额': df['投入金额'][i],
                        '数量': df['数量'][i],
                        '单价': df['单价'][i],
                        '金额': df['金额'][i]
                        }
            # 出现新的编号
            if dict_data.get(str_cp_no) is None:
                # dict_son = {}

                dict_son = {str_wl_no: list_row}
                # dict_son = {}
                # dict_son[str_wl_no] = list_row
                dict_data[str_cp_no] = dict_son
            else:
                # print(dict_data)
                dict_data[str_cp_no][str_wl_no] = list_row
    # print("-----------------------------------------------")
    # print(dict_data['101000105']['704015094'])
    # print(dict_data['101000124']['802900023'])

    # print(df['产品编号'])
    return dict_data


if __name__ == '__main__':
    dict_a = extract_20196()
    print(dict_a)




