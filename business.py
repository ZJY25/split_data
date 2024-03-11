import openpyxl
business_type = ["回网",
                 "预离",
                 "全能卡",
                 "88", "118", "158", "188", "228", "288", "388", "39", "79", "119", "229",
                 "单移", "双百", "副卡"]
net_back_key_words = ["互", "宽", "高", "有效", "标清"]


def judge_business_type(business_str):
    for key, value in enumerate(business_type):
        flag = business_str.find(value)
        if flag >= 0:
            print(business_str + " " + value + " " + str(key))
            return key

def judge_net_back_type(net_back_str):
    for key, value in enumerate(net_back_key_words):
        flag = net_back_str.find(value)
        if flag >= 0:
            print(business_str + " " + value + " " + str(key))
            return key

def business_process(business_str, excel, i, j):
    b_type = judge_business_type(business_str)
    print("b_type: " + str(b_type))
    if b_type == 0:
        print("回网")
        if [0] in business_str:
            excel.write_cell(i, j, business_str)
            print("写入excel")
    if b_type == 1:
        print("预离")
    if b_type == 2:
        print("全能卡")
    if b_type in range(3, 14):
        print("数字新增")
    if b_type in range(14, 17):
        print("文字新增")


if __name__ == '__main__':
    business_process("副卡")
