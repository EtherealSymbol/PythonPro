import xlrd


if __name__ == '__main__':
    book = xlrd.open_workbook('表.xls')
    names = book.sheet_names()
    print(names)
    first_sheet = book.sheet_by_index(0)














