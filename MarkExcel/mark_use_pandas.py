import pandas as pd
import re
import xlwt
import xlrd
from xlutils.copy import copy

from data_process import extract_num


if __name__ == '__main__':
    file_name = '表 - 副本.xls'
    df = pd.read_excel(file_name)

    rb = xlrd.open_workbook(file_name)
    sh = rb.sheets()[0]
    wb = copy(rb)
    ws = wb.get_sheet(0)

    # print(df)
    # print(df.values)
    # print(df['产品规格'])
    # print(df['物料'])
    # print(df['数量'])
    # print(df.columns)
    # # 初始化物料，产品规格，数量为-1
    # wl = -1
    # cp = -1
    # sl = -1
    k = 0
    for elem in df.columns:
        # 获得“物料”所在列
        if elem == "物料":
            wl = k
        elif elem == "产品规格":
            cp = k
        elif elem == "数量":
            sl = k
        k += 1
    print(cp, wl, sl)
    # 正则
    reg_ab = r"安瓿"
    reg_ch = r"彩盒"
    reg_sms = r"说明书"
    reg_tp = r"托盘"
    # 定义i获取row之后的行
    i = 0
    all_rows = df.values
    # row是当前行
    for row in all_rows:
        # print(row)
        # 获取每一行物料数据
        wl_data = row[wl]
        # 获取该物料对应数量
        sl_num = row[sl]
        # 是安瓿
        if wl_data is not None and re.search(reg_ab, str(wl_data)):
            # 产品规格字符串
            cp_str = row[cp]
            # 从’产品规格‘中提取‘数量规格’
            cp_num = extract_num(cp_str)
            # print(wl_data, sl_num, cp_str, cp_num)
            # 判断后面连续的是否也为“安瓿”
            j = 1
            next_wl_data = all_rows[i+1][wl]
            # print(wl_data, next_wl_data)
            while re.search(reg_ab, str(next_wl_data)):
                i += 1
                # 数量相加
                sl_num += row[sl]
                next_wl_data = all_rows[i+j][wl]
                j += 1
                # print("还是安瓿")
            # 下面判断数据是否符合要求，误差在0.5内
            if float(sl_num) < float(cp_num) or float(sl_num) - float(cp_num) > 0.5:
                # 处理
                # print(cp_str, cp_num, sl_num, "数据不合格", j)
                # 将数据不合格这几行数据标红
                for r in range(j):
                    row_num = i - j + r + 1
                    print("这条不合格！！！！", r, row_num)
                    # print(df.loc[row_num][sl])
                    print(df.iloc[row_num, sl])
                    # 修改样式
                    # mark(file_name, row_num, sl, df.iloc[row_num, sl])
                    # def highlight_max():
                    #     print(11)
                    #     return pd.DataFrame(index=1, columns=1)
                    # df.style.apply(highlight_max, color='darkorange', axis=None)
                    # print(df.style)
                    red_style = xlwt.easyxf('pattern: fore_colour red;')
                    ws.write(row_num, sl, df.iloc[row_num, sl], red_style)


        # else:

        i += 1
    # df.iloc[1, 1] = "WWWWWW"
    # df.style.highlight_null('red')
    # df.to_excel(file_name, index=False, header=True)
    wb.save(file_name)






