import re
import excel_process

import openpyxl
business_type = ["预离",
                 "回网",
                 "全能卡",
                 "39", "118", "158", "188", "228", "288", "388", "88", "79", "119", "229",
                 "单移", "双百", "副卡",
                 "新增",
                 "新开",
                 ]
net_back_key_words = ["互", "宽", "高", "有效", "标清"]

all_can_key_words = ["99", "129", "169", "229"]


def judge_business_type(business_str):
    for key, value in enumerate(business_type):
        flag = business_str.find(value)
        if flag >= 0:
            print(business_str + " " + value + " " + str(key))
            return key


def judge_ghk_type(net_back_str):
    temp_ghk = [0, 0, 0, 0, 0]
    for key, value in enumerate(net_back_key_words):
        flag = net_back_str.find(value)
        if flag >= 0:
            temp_ghk[key] = 1
    return temp_ghk


def business_process(business_str, excel, i, j):
    b_type = judge_business_type(business_str)
    print("b_type: " + str(b_type))
    if b_type == 1:
        print("回网")
        ghk = judge_ghk_type(business_str)
        for key, value in enumerate(ghk):
            print(str(key) + ":" + str(value))
            if value > 0:
                if key == 0:
                    excel.write_cell(i, "M", "互动回网")
                if key == 1:
                    excel.write_cell(i, "N", "宽带回网")
                if key in range(2, 5):
                    excel.write_cell(i, "L", "有效回网")
    if b_type == 0:
        print("预离")
        ghk = judge_ghk_type(business_str)
        judge_ghk_type(business_str)
        for key, value in enumerate(ghk):
            if value > 0:
                if key == 0:
                    excel.write_cell(i, "Q", "毛都没有")
                if key == 1:
                    excel.write_cell(i, "P", "宽带预离网回网")
                if key in range(2, 5):
                    excel.write_cell(i, "O", "有效预离网回网")
    if b_type == 2:
        print("全能卡")
        sup_card = re.findall("副", business_str)
        temp_number = re.findall("\d+", business_str)
        the_temp_number = temp_number[0]
        excel.write_interactive_add(i)
        excel.write_valid_add(i)
        excel.write_broadband_add(i)
        print(the_temp_number)
        for key, value in enumerate(all_can_key_words):
            flag = the_temp_number.find(value)
            if flag >= 0:
                if key == 0:
                    excel.write_G(i, "全能99")
                if key == 1:
                    excel.write_G(i, "a")
                if key == 2:
                    excel.write_G(i, "b")
                if key == 3:
                    excel.write_G(i, "MAX")
        if sup_card:
            excel.write_H(i, "单移/副卡")

    if b_type in range(3, 14):
        number = re.findall("\d+", business_str)
        sup_card = re.findall("副", business_str)
        print(sup_card)
        the_number = number[0]
        print("数字新增")
        if b_type == 3:
            excel.write_G(i, the_number)
            excel.write_valid_add(i)
        if b_type in range(4, 14):
            excel.write_G(i, the_number)
            excel.write_interactive_add(i)
            excel.write_valid_add(i)
            excel.write_broadband_add(i)
        if sup_card:
            excel.write_H(i, "单移/副卡")

    if b_type in range(14, 17):
        print("文字新增")
        excel.write_H(i, "单移/副卡")

    if b_type in range(17, 19):
        print("高互宽新增")
        ghk = judge_ghk_type(business_str)
        for key, value in enumerate(ghk):
            if value > 0:
                if key == 0:
                    excel.write_interactive_add(i)
                if key == 1:
                    excel.write_broadband_add(i)
                if key in range(2, 5):
                    excel.write_valid_add(i)
