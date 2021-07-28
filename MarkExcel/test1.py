import os
import re
import openpyxl
# from data_process import extract_num, handle_tag
from openpyxl.styles import colors, PatternFill
import pandas as pd
import win32com.client as win32

if __name__ == '__main__':
    # f = '表 - 副本.xls'
    # file_name_be, suff = os.path.splitext(f)  # 对路径进行分割，分别为文件路径和文件后缀
    # if suff == '.xls':
    #     print('将对{}文件进行转换...'.format(f))
    #     data = pd.DataFrame(pd.read_excel(f))  # 读取xls文件
    #     data.to_excel(file_name_be + '格式转变.xlsx', index=False)  # 格式转换
    #     # print(' {} 文件已转化为 {} 保存在 {} 目录下\n'.format(f, file_name_be + '格式转变.xlsx'))
    fname = r"C:\Users\ty\Desktop\WorkPlace\MarkExcel\表2.xls"
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)

    wb.SaveAs(fname + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
    wb.Close()  # FileFormat = 56 is for .xls extension
    excel.Application.Quit()


