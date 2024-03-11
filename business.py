business_type = ["回网", "预离网回网",
                 "全能卡"
                 "88", "118", "158", "188", "228", "288", "388", "39", "79", "119", "229",
                 "单移", "双百", "副卡"]


def judge_business_type(business_str):
    for key, value in enumerate(business_type):
        if business_str.find(value):
            print(business_str.find(value))
            print(business_str + " " + value + " " + str(key))
            return key


def business_process(business_str):
    b_type = judge_business_type(business_str)
    print(b_type)
    if b_type in (0, 1):
        print("回网")


if __name__ == '__main__':
    business_process("回网高互宽")
