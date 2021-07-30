import openpyxl


def extract_second():
    wb = openpyxl.load_workbook("产品成本计算单查询 - 副本.xlsx", data_only=True)
    # active获得的是sheet1
    # ws = wb.active
    ws = wb['Sheet']
    # 获得所有列
    # for cols in ws.columns:
    #     for col in cols:
    #         print(col.value)
    # 获得第一行
    # print(list(rows))
    rows = ws.rows
    list_rows = list(rows)
    first_row = list_rows[0]
    # for elem in first_row:
    #     print(elem.value)
    #     if elem.value == "物料":
    # print(len(first_row))
    first_row_len = len(first_row)
    # 逆序遍历，从后往前找物料等信息，因为他们有两列
    # 分别记录他们在第几列，并将其初始化为-1
    s_wl = -1
    s_gg = -1
    s_dw = -1
    s_tr_sl = -1
    s_tr_je = -1
    s_sl = -1
    s_dj = -1
    s_je = -1
    for i in range(first_row_len, 1, -1):
        if first_row[i-1].value == '物料' and s_wl == -1:
            s_wl = i
        elif first_row[i-1].value == '物料规格' and s_gg == -1:
            s_gg = i
        elif first_row[i-1].value == '物料计量单位' and s_dw == -1:
            s_dw = i
        elif first_row[i-1].value == '投入数量' and s_tr_sl == -1:
            s_tr_sl = i
        elif first_row[i-1].value == '投入金额' and s_tr_je == -1:
            s_tr_je = i
        elif first_row[i-1].value == '数量' and s_sl == -1:
            s_sl = i
        elif first_row[i-1].value == '单价' and s_dj == -1:
            s_dj = i
        elif first_row[i-1].value == '金额' and s_je == -1:
            s_je = i
    # print(s_wl, s_gg, s_dw, s_tr_sl, s_tr_je, s_sl, s_dj, s_je)
    dict_second = {'s_wl': s_wl,
                   's_gg': s_gg,
                   's_dw': s_dw,
                   's_tr_sl': s_tr_sl,
                   's_tr_je': s_tr_je,
                   's_sl': s_sl,
                   's_dj': s_dj,
                   's_je': s_je}

    wb.close()
    return dict_second



