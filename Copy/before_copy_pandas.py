import pandas as pd


if __name__ == '__main__':
    file_name = '产品成本计算单查询 - 副本.xlsx'
    df = pd.read_excel(file_name, engine='openpyxl')
    # print(df)
    print(df['物料.1'][0])
    df['物料.1'][2] = 1212212
    df.to_excel('new.xlsx')


