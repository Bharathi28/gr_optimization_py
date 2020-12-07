import xlrd

def getExcelData(input_path, sheet_name):

    book = xlrd.open_workbook(input_path)
    sheet = book.sheet_by_name(sheet_name)

    # data = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    # print(data)
    # return data

    for i in range(1, sheet.nrows):
        print(i)

