import re
import business as bp
from excel_process import ExcelProcess


path = r"D:\projects\split_data"
filename = "test.xlsx"
table_name = "test"
name_pattern = "发展人[:：,，][（）\d\u4e00-\u9fa5]+"
date_pattern = "\d+月\d+日"
business_pattern = "(?<=业务[:：])[\d\u4e00-\u9fa5]+(?=[,，])"

excel = ExcelProcess(path, filename)
excel.load_excel_sheet(table_name)
for row in range(5, 10):

    column = 2
    all_string = excel.get_cell(row, 1)

    names = re.findall(name_pattern, all_string)
    for name in names:
        name = re.findall("(?<=[:：,，]).*$", name)
        name_str = name[0]
        excel.write_cell(row, column, name_str)
        column += 1

    column = 4

    date = re.findall(date_pattern, all_string)
    date_str = date[0]
    excel.write_cell(row, column, date_str)
    column += 1

    business = re.findall(business_pattern, all_string)
    business_str = business[0]
    excel.write_cell(row, column, business_str)
    column += 1
    bp.business_process(business_str, excel, row, column)

excel.save_excel(filename)
