import re
import business as bp
from excel_process import ExcelProcess


path = r"D:\projects\split_data"
filename = "自动化三月红.xlsx"
table_name = "号卡固网晒单"
station_pattern = "(?<=日)[\u4e00-\u9fa5]+(?=[第])"
business_pattern = "(?<=业务[:：])[＋（）\d\u4e00-\u9fa5]+(?=[,，. 。：])"

excel = ExcelProcess(path, filename)
excel.load_excel_sheet(table_name)

for row in range(2, 12):
    print(row)
    all_string = excel.get_cell(row, 2)
    print(all_string)

    station = re.findall(station_pattern, all_string)
    print(station)
    excel.write_cell(row, "C", station[0])

    business = re.findall(business_pattern, all_string)[0]
    print(business)

    card_type = bp.judge_card(business_str=business)
    print(card_type)

    for key, value in enumerate(card_type):
        if value > 0:
            if key == 0:
                excel.write_cell(row, "F", "全能")
            if key == 1:
                excel.write_cell(row, "D", "卡")
            if key == 2:
                excel.write_cell(row, "E", "副卡")

excel.save_excel(filename)