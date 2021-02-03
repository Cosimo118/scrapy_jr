import xlrd
import xlwt
from xlutils.copy import copy

def get_error_list(path,path_2):
    data = xlrd.open_workbook(path)
    table = data.sheet_by_name('My Worksheet')
    name = table.name
    rowNum = table.nrows
    lineNum = table.ncols
    print(lineNum)
    workbook = copy(xlrd.open_workbook(path_2))
    worksheet = workbook.get_sheet(0)
    for i in range(2,rowNum):
        if table.cell(i,34).ctype is not 0:
            td_id = int(table.cell(i,34).value)
            if td_id < 60000:
                print(td_id)
                for j in range(0,lineNum):
                    td_value = table.cell(i,j).value
                    worksheet.write(i,j,label=str(td_value))
    workbook.save(path_2)

path = r'C:\Users\Administrator\Desktop\vscode_workspace\scrapy\tutorial\tutorial\Excel_02.xls'
path_2 = r'C:\Users\Administrator\Desktop\vscode_workspace\scrapy\tutorial\tutorial\Excel_01_copy.xls'

get_error_list(path,path_2)