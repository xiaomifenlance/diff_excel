#!/usr/bin/env python
# -*- coding: UTF-8 -*-
# Author:Lzh
import xlrd
import time
def read_excel(ori_path, tar_path, ori_sheet_nums, tar_sheet_nums):

    diff_dic = []  #创建一个保存不同内容的列表，备用
    ori_data = xlrd.open_workbook(ori_path)  #用xlrd模块打开源和目标excel表格
    tar_data = xlrd.open_workbook(tar_path)

    # ori_nums = len(ori_data) #求出excel内有几张表格
    # tar_nums = len(tar_data)

    print("读取源表格...打开第%s张表格\n"
          "读取目标表格...打开第%s张表格" % (ori_sheet_nums+1, tar_sheet_nums+1))
    ori_sheet_name = ori_data.sheets()[ori_sheet_nums]  #获取源表格的内存地址
    print("获取当地时间...ok")
    log_name = "Result_%s.txt" % (time.strftime("%y_%m_%d_%H_%M_%S")) #输入结果的文件名

    f = open(log_name, 'w') #打开输出结果文件，如果不存在则创建
    print("打开创建输出文件: %s...ok\n####################################################\n\n The Program Starting..." % log_name)

    tar_sheet_name = tar_data.sheets()[tar_sheet_nums] #获取目标表格内存地址

    ori_nrow = ori_sheet_name.nrows #获取行数
    tar_nrow = tar_sheet_name.nrows


    for i in range(ori_nrow): #通过对比每一行的字典 判断是否重复
        for k in range(tar_nrow):
            if ori_sheet_name.row_values(i) == tar_sheet_name.row_values(k):
                # print("%s 的第 %d 行与 %s 的第 %d 行 重复：" % (ori_path, i, tar_path, k))
                # print("%s == %s\n\n" % (ori_sheet_name.row_values(i), tar_sheet_name.row_values(k)))
                f.write("%s 的第 %d 行与 %s 的第 %d 行 重复：\n" % (ori_path, i, tar_path, k))
                f.write("%s == %s\n\n" % (ori_sheet_name.row_values(i), tar_sheet_name.row_values(k)))
                if i + 1 not in diff_dic:
                    diff_dic.append(i + 1)

    print("重复的行数如下：", diff_dic, "\n")
    print("详情请查看文件: %s...ok\n####################################################\n" % log_name)
    f.write("重复的行数如下：%s" % str(diff_dic))
    f.close()  # 关闭文件
    print("Done\n####################################################\n")  # 程序完成标志




Describe_program = """
Compare for Excel 1.0
###########################################################################
    Author: Lance
    Time: Nov. 4, 2018
    Describe:
    The tool can find differents between two excels by comparing every row.
###########################################################################"""
print(Describe_program)
ori_file_name = input("请输入源文件的路径:")
ori_sheet_nums = input("需要比对的表格编号（输入1则为第一张表，默认为第一张表）:")

tar_file_name = input("请输入目标文件的路径:")
tar_sheet_nums = input("需要比对的表格编号（输入1则为第一张表，默认为第一张表）:")


ori_nums = int(ori_sheet_nums)-1
tar_nums = int(tar_sheet_nums)-1

read_excel(ori_file_name, tar_file_name, ori_nums, tar_nums)

while True:
    continue_flag = input("是否继续比较目前的两个excel: '%s' 和 '%s' (x/y)?:" %(ori_file_name,tar_file_name))
    if continue_flag == 'x' or continue_flag == 'X':
        print('\nThe Program End.')
        break
    if continue_flag == 'y' or continue_flag == 'y':
        ori_sheet_nums = input("输入 '%s' 需要比对的表格编号:" %(ori_file_name))
        tar_sheet_nums = input("输入 '%s' 需要比对的表格编号:" %(tar_file_name))

        ori_nums = int(ori_sheet_nums) - 1
        tar_nums = int(tar_sheet_nums) - 1

        read_excel(ori_file_name, tar_file_name, ori_nums, tar_nums)

    else:
        print("\n输入异常：请输入x退出或y继续比对...\n")
