import openpyxl
business_type = ["预离",
                 "回网",
                 "全能卡",
                 "88", "118", "158", "188", "228", "288", "388", "39", "79", "119", "229",
                 "单移", "双百", "副卡"]
net_back_key_words = ["互", "宽", "高", "有效", "标清"]

temp_net_back = [0, 0, 0, 0, 0]


def judge_business_type(business_str):
    for key, value in enumerate(business_type):
        flag = business_str.find(value)
        if flag >= 0:
            print(business_str + " " + value + " " + str(key))
            return key


def judge_detail_type(net_back_str):
    for key, value in enumerate(net_back_key_words):
        flag = net_back_str.find(value)
        if flag >= 0:
            temp_net_back[key] = 1


def business_process(business_str, excel, i, j):
    b_type = judge_business_type(business_str)
    print("b_type: " + str(b_type))
    if b_type == 1:
        print("回网")
        judge_detail_type(business_str)
        for key, value in enumerate(temp_net_back):
            if value > 0:
                if key == 0:
                    excel.write_cell(i, 12, "互动回网")
                if key == 1:
                    excel.write_cell(i, 13, "宽带回网")
                if key in range(2, 5):
                    excel.write_cell(i, 11, "有效回网")
    if b_type == 0:
        print("预离")
        judge_detail_type(business_str)
        for key, value in enumerate(temp_net_back):
            if value > 0:
                if key == 0:
                    excel.write_cell(i, 12, "毛都没有")
                if key == 1:
                    excel.write_cell(i, 13, "宽带预离网回网")
                if key in range(2, 5):
                    excel.write_cell(i, 11, "有效预离网回网")
    if b_type == 2:
        print("全能卡")
    if b_type in range(3, 14):
        print("数字新增")
    if b_type in range(14, 17):
        print("文字新增")


