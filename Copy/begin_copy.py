import openpyxl
from before_copy import extract_second
from extract_data_to_dict_2021 import extract_data_to_dict_fun
from Cost_in2021.extract_data_to_dict_20215 import extract_20215

if __name__ == '__main__':
    file_name = "产品成本计算单查询 - 副本.xlsx"
    wb = openpyxl.load_workbook(file_name, data_only=True)

    ws = wb['Sheet']

    rows = ws.rows
    list_rows = list(rows)
    # for li in list_rows:
    #     for n in li:
    #         print(n.value, end='  ')
    #     print()
    first_row = list_rows[0]
    # for f in first_row:
    #     print(f.value)

    first_row_len = len(first_row)
    # 逆序遍历，从后往前找物料等信息，因为他们有两列
    # 分别记录他们在第几列，并将其初始化为-1
    cp_no = -1  # 产品编号
    f_wl_no = -1  # 物料编号
    f_wl = -1
    f_gg = -1
    f_dw = -1
    f_tr_sl = -1
    f_tr_je = -1
    f_sl = -1
    f_dj = -1
    f_je = -1
    for i in range(1, first_row_len):
        if first_row[i - 1].value == '产品编号' and cp_no == -1:
            cp_no = i
        if first_row[i - 1].value == '物料编号' and f_wl_no == -1:
            f_wl_no = i
        elif first_row[i - 1].value == '物料' and f_wl == -1:
            f_wl = i
        elif first_row[i - 1].value == '物料规格' and f_gg == -1:
            f_gg = i
        elif first_row[i - 1].value == '物料计量单位' and f_dw == -1:
            f_dw = i
        elif first_row[i - 1].value == '投入数量' and f_tr_sl == -1:
            f_tr_sl = i
        elif first_row[i - 1].value == '投入金额' and f_tr_je == -1:
            f_tr_je = i
        elif first_row[i - 1].value == '数量' and f_sl == -1:
            f_sl = i
        elif first_row[i - 1].value == '单价' and f_dj == -1:
            f_dj = i
        elif first_row[i - 1].value == '金额' and f_je == -1:
            f_je = i
    # print(f_wl_no, f_wl, f_gg, f_dw, f_tr_sl, f_tr_je, f_sl, f_dj, f_je)
    dict_sec = extract_second()

    dict_compare = extract_data_to_dict_fun("2021年07月小针成本指标表.xls", "2021.07与预算")

    for li in list_rows[1:]:
        cp_no_data = li[cp_no-1].value
        f_wl_no_data = li[f_wl_no-1].value
        f_wl_data = li[f_wl-1].value
        s_wl_data = li[dict_sec['s_wl']-1].value
        if f_wl_no_data is not None and f_wl_data is not None and s_wl_data is None:
            # 开始寻找数据并复制数据
            month_dy = dict_compare.get(cp_no_data)
            print(f_wl_no_data, f_wl_data, s_wl_data, month_dy)



    wb.close()




