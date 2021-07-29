import numpy as np
import pandas as pd


def recent_from_7_to_1(list_row):
    # 7月在第11列，6月在第10列，以此类推，5月9列，4月8列，3月7列，2月6列，1月5列
    # 反应到下标应该-1
    # 规定，如果没有找到最近的月份，返回-1
    latest = -1
    # range(10, 3)有10，9，8，7，6，5，4
    for j in range(10, 3, -1):
        per_month = list_row[j]
        # print(list_row[j])
        # print(j)
        if 1 - np.isnan(per_month):
            # 11-4=4, 10-6=4有规律10+1-4=7月
            latest = j+1-4
            # print(latest)
            break
    return latest


def extract_data_to_dict_fun(file_name, the_sheet_name):
    # file_name = '2021年07月小针成本指标表.xls'
    # df = pd.read_excel(file_name, sheet_name='2021.07与预算')
    df = pd.read_excel(file_name, sheet_name=the_sheet_name)
    # 获得表头, 表头在第20行
    title = df.values[18]
    # print(title)
    # print(df.values[1:])
    # print(df.head(20))
    df_values = df.values
    len_values = len(df_values)

    # 将表格中所有数据存入一个字典，一个编号对应一个月份
    dict_all = {}

    for i in range(19, len_values):
        row_value = df_values[i]
        # 7月在第10列，6月在第9列，以此类推，5月8列，4月7列，3月6列，2月5列，1月4列
        # print(row_value, row_value[9])
        '''
        row_value: [1 '0101000293' '地塞米松磷酸钠注射液' '1ml:2mg*10支*300盒' nan nan nan nan nan 0.8
        nan nan nan nan nan nan 0.8 0.7799999999999999 0.75 0.7873384758364312
        0.61 nan nan nan nan nan 1 '0101000293' '地塞米松磷酸钠注射液' '1ml:2mg*10支*300盒'
        nan nan 0.73 0.69 nan 0.86 nan 0.78 nan 0.84 nan nan 0.7799999999999999
        0.75 0.7873384758364312 0.61]
        '''
        # 产品代码
        cp_dm = row_value[1]
        # 保证不是nan
        if 1 - pd.isna(cp_dm):
            # print(cp_dm)
            latest_data = recent_from_7_to_1(row_value)
            # 如果遇到之前出现过，再比较是不是最近的
            if dict_all.get(cp_dm) is None:
                dict_all[cp_dm] = latest_data
            else:
                if dict_all[cp_dm] > latest_data:
                    dict_all[cp_dm] = latest_data

    # for key in dict_all:
    #     print(key, dict_all[key])
    '''
    字典数据提取完毕
    '''
    return dict_all


# if __name__ == '__main__':
    # dict_compare = extract_data_to_dict_fun("2021年07月小针成本指标表.xls", "2021.07与预算")
    # print(dict_compare)
    # for key in dict_compare:
        # print(key, dict_compare[key])


