import re
import business as bp
from excel_process import ExcelProcess


path = r"D:\projects\split_data\split_data"
filename = "自动化三月红.xlsx"
table_name = "号卡固网晒单"
name_pattern = "发展人[:：,，][（）\d\u4e00-\u9fa5]+"
date_pattern = "\d+月\d+日"
business_pattern = "(?<=业务[:：])[（）\d\u4e00-\u9fa5]+(?=[,， 。：])"
xjk = ["王烨", "刘逢贵", "林灿光"]

excel = ExcelProcess(path, filename)
excel.load_excel_sheet(table_name)
print(excel.get_max_row())
for row in range(2, 34):
    print(str(row) + ":")
    column = 3
    all_string = excel.get_cell(row, 2)

    names = re.findall(name_pattern, all_string)
    for name in names:
        name = re.findall("(?<=[:：,，]).*$", name)
        name_str = name[0]
        for x_name in xjk:
            if re.findall(x_name, name_str):
                name_str = "小集客"
        excel.write_cell(row, column, name_str)
        column += 1

    column = 4

    date = re.findall(date_pattern, all_string)
    date_str = date[0]
    excel.write_cell(row, "A", date_str)
    column += 1

    business = re.findall(business_pattern, all_string)
    business_str = business[0]
    excel.write_cell(row, "E", business_str)
    column += 1
    bp.business_process(business_str, excel, row, column)

excel.save_excel(filename)
