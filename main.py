import os
import openpyxl
from openpyxl.utils import get_column_letter
import re

path = r"D:\projects\pythonProject"
os.chdir(path)  # 修改工作路径

workbook = openpyxl.load_workbook('test.xlsx')
# print(workbook.sheetnames)  # 查看work book的表格

sheet = workbook['test']
# print(sheet)

max_row = sheet.max_row
print("max_row", max_row)  # 获取最大行数

column_item = sheet['A1':'A%d' % max_row]
print("column_item", column_item)

for i in range(23, 27):
    column = 2

    all_string = sheet.cell(row=i, column=1).value
    print(str(i), ":", all_string)

    date = re.findall('\d+月\d+日', all_string)
    date_str = date[0]
    # print(date_str)
    column_letter = get_column_letter(column)
    # print(column_letter + str(i))
    sheet[column_letter + str(i)] = date_str
    column += 1

    business = re.findall('(?<=业务[:：])[\d\u4e00-\u9fa5]+(?=[,，])', all_string)
    business_str = business[0]
    # print(business_str)
    column_letter = get_column_letter(column)
    # print(column_letter + str(i))
    sheet[column_letter + str(i)] = business_str
    column += 1

    # ((\d*) | ([(（][\u4e00-\u9fa5]+[)）])*)
    names = re.findall('发展人[:：,，][\u4e00-\u9fa5]+', all_string)
    print(names)
    for name in names:
        name = re.findall("(?<=[:：,，]).*$", name)
        print(name)
        name_str = name[0]
        column_letter = get_column_letter(column)
        print(column_letter + str(i))
        sheet[column_letter + str(i)] = name_str
        column += 1

# workbook.save('test.xlsx')