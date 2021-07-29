import numpy as np
import pandas as pd

if __name__ == '__main__':
    file_name = '2021年6月产品成本计算单查询.xls'
    df = pd.read_excel(file_name)

    # print(df.values)
    # for row in df.values:
    #     print(row)

    # 将所有数据提取出来存到字典中
    dict_data = {}

    '''
    数据结构：
    dict_data:{
        dict_sno: {
            list_row: [],
            list_row: [],
            ...
        },
        dict_sno: {
            list_row: [],
            list_row: [],
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
        print(df_values[i])
        print(df['产品编号'][i], df['物料'][i])
        # 布尔取反
        row_cp_bh = df['产品编号'][i]
        if 1 - np.isnan(row_cp_bh):
            # row是101000338.0这类的数字
            # 转化成字符串并去掉后面的.0
            str_cp_no = str(row_cp_bh)[0:-1]
            # print(str_cp_no)
            # 出现新的编号
            if dict_data.get(str_cp_no) is not None:
                dict_son = {}
                list_row = []

    # print(df['产品编号'])



