#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com


import xlsxwriter
import os
import xlrd
import time
import datetime
from threading import Thread
import wx
import matplotlib.pyplot as pyplot
import numpy
from collections import Counter


filename_original_zonglan = "测试验证部专利总览包含各节点时间-20171102.xlsx".decode('gbk')
file_name = xlrd.open_workbook(filename_original_zonglan, encoding_override='cp936')
sheet_filter_one = file_name.sheet_by_index(0)
total_rows_one = sheet_filter_one.nrows
list_date_tijiao_start = []
list_date_tijiao_jiekouren  = []
list_date_tijiao_lingdao = []
list_date_tijiao_final = []
list_date_zhuanxie_start = []
list_date_zhuanxie_daili = []
list_date_zhuanxie_confirm = []
list_date_zhuanxie_final = []

for line_count in range(2, total_rows_one):
    date_tijiao_start = sheet_filter_one.cell(line_count, 10).value.strip()
    date_tijiao_jiekouren = sheet_filter_one.cell(line_count, 11).value.strip()
    date_tijiao_lingdao = sheet_filter_one.cell(line_count, 12).value.strip()
    date_tijiao_final = sheet_filter_one.cell(line_count, 13).value.strip()
    date_zhuanxie_start = sheet_filter_one.cell(line_count, 14).value.strip()
    date_zhuanxie_daili = sheet_filter_one.cell(line_count, 15).value.strip()
    date_zhuanxie_confirm = sheet_filter_one.cell(line_count, 16).value.strip()
    date_zhuanxie_final = sheet_filter_one.cell(line_count, 17).value.strip()
    list_date_tijiao_start.append(date_tijiao_start)
    list_date_tijiao_jiekouren.append(date_tijiao_jiekouren)
    list_date_tijiao_lingdao.append(date_tijiao_lingdao)
    list_date_tijiao_final.append(date_tijiao_final)
    list_date_zhuanxie_start.append(date_zhuanxie_start)
    list_date_zhuanxie_daili.append(date_zhuanxie_daili)
    list_date_zhuanxie_confirm.append(date_zhuanxie_confirm)
    list_date_zhuanxie_final.append(date_zhuanxie_final)

list_days_tijiao_jiekouren = []
list_days_tijiao_lingdao = []
list_days_tijiao_final = []
list_days_zhuanxie_start = []
list_days_zhuanxie_daili = []
list_days_zhuanxie_confirm = []
list_days_zhuanxie_final = []
now_date = datetime.datetime.now()
now_year = now_date.strftime("%Y")
now_month = now_date.strftime("%m")
now_days = now_date.strftime("%d")

for index_data, item_data in enumerate(list_date_tijiao_start):
    #发明人提交日期为None时，全部时间为None
    if item_data == "None":
        continue
    else:
        #获取接口人处理时间
        #如果接口人日期为None，则根据提交和现在时间确定
        if list_date_tijiao_jiekouren[index_data] == "None":
            days_tijiao_jiekouren = (
            datetime.datetime(int(now_year), int(now_month), int(now_days)) - datetime.datetime(
                int(item_data.split("-")[0]), int(item_data.split("-")[1]), int(item_data.split("-")[2]))).days
            days_tijiao_lingdao = "None"
            days_tijiao_final = "None"
            days_zhuanxie_start = "None"
            days_zhuanxie_daili = "None"
            days_zhuanxie_confirm = "None"
            days_zhuanxie_final = "None"
        else:
            #如果接口人日期不为None，则接口人日期减去发明人提交日期
            days_tijiao_jiekouren = (datetime.datetime(int(list_date_tijiao_jiekouren[index_data].split("-")[0]),
                                                       int(list_date_tijiao_jiekouren[index_data].split("-")[1]), int(
                    list_date_tijiao_jiekouren[index_data].split("-")[2])) - datetime.datetime(
                int(item_data.split("-")[0]), int(item_data.split("-")[1]), int(item_data.split("-")[2]))).days
            #获取领导审批时间
            #领导审批日期为None时
            if list_date_tijiao_lingdao[index_data] == "None":
                days_tijiao_lingdao = (
                datetime.datetime(int(now_year), int(now_month), int(now_days)) - datetime.datetime(
                    int(list_date_tijiao_jiekouren[index_data].split("-")[0]),
                    int(list_date_tijiao_jiekouren[index_data].split("-")[1]),
                    int(list_date_tijiao_jiekouren[index_data].split("-")[2]))).days
                days_tijiao_final = "None"
                days_zhuanxie_start = "None"
                days_zhuanxie_daili = "None"
                days_zhuanxie_confirm = "None"
                days_zhuanxie_final = "None"
            else:
                #领导日期不为None时
                days_tijiao_lingdao = (datetime.datetime(int(list_date_tijiao_lingdao[index_data].split("-")[0]),
                                                         int(list_date_tijiao_lingdao[index_data].split("-")[1]), int(
                        list_date_tijiao_lingdao[index_data].split("-")[2])) - datetime.datetime(
                    int(list_date_tijiao_jiekouren[index_data].split("-")[0]),
                    int(list_date_tijiao_jiekouren[index_data].split("-")[1]),
                    int(list_date_tijiao_jiekouren[index_data].split("-")[2]))).days
                #获取专利工程师确认时间
                #专利工程师确认日期为None时
                if list_date_tijiao_final[index_data] == "None":
                    days_tijiao_final = (
                    datetime.datetime(int(now_year), int(now_month), int(now_days)) - datetime.datetime(
                        int(list_date_tijiao_lingdao[index_data].split("-")[0]),
                        int(list_date_tijiao_lingdao[index_data].split("-")[1]),
                        int(list_date_tijiao_lingdao[index_data].split("-")[2]))).days
                    days_zhuanxie_start = "None"
                    days_zhuanxie_daili = "None"
                    days_zhuanxie_confirm = "None"
                    days_zhuanxie_final = "None"
                else:
                    #专利工程师确认日期不为None时
                    days_tijiao_final = (
                        datetime.datetime(int(list_date_tijiao_final[index_data].split("-")[0]),
                                          int(list_date_tijiao_final[index_data].split("-")[1]),
                                          int(list_date_tijiao_final[index_data].split("-")[2])) - datetime.datetime(
                            int(list_date_tijiao_lingdao[index_data].split("-")[0]),
                            int(list_date_tijiao_lingdao[index_data].split("-")[1]),
                            int(list_date_tijiao_lingdao[index_data].split("-")[2]))).days
                    if list_date_zhuanxie_start[index_data] == "None":
                        days_zhuanxie_start = (
                            datetime.datetime(int(now_year), int(now_month), int(now_days)) - datetime.datetime(
                                int(list_date_tijiao_final[index_data].split("-")[0]),
                                int(list_date_tijiao_final[index_data].split("-")[1]),
                                int(list_date_tijiao_final[index_data].split("-")[2]))).days
                        days_zhuanxie_daili = "None"
                        days_zhuanxie_confirm = "None"
                        days_zhuanxie_final = "None"
                    else:
                        days_zhuanxie_start = (
                            datetime.datetime(int(list_date_zhuanxie_start[index_data].split("-")[0]),
                                              int(list_date_zhuanxie_start[index_data].split("-")[1]),
                                              int(list_date_zhuanxie_start[index_data].split("-")[2])) - datetime.datetime(
                                int(list_date_tijiao_final[index_data].split("-")[0]),
                                int(list_date_tijiao_final[index_data].split("-")[1]),
                                int(list_date_tijiao_final[index_data].split("-")[2]))).days
                        if list_date_zhuanxie_daili[index_data] == "None":
                            days_zhuanxie_daili = (
                                datetime.datetime(int(now_year), int(now_month), int(now_days)) - datetime.datetime(
                                    int(list_date_zhuanxie_start[index_data].split("-")[0]),
                                    int(list_date_zhuanxie_start[index_data].split("-")[1]),
                                    int(list_date_zhuanxie_start[index_data].split("-")[2]))).days
                            days_zhuanxie_confirm = "None"
                            days_zhuanxie_final = "None"
                        else:
                            days_zhuanxie_daili = (
                                datetime.datetime(int(list_date_zhuanxie_daili[index_data].split("-")[0]),
                                                  int(list_date_zhuanxie_daili[index_data].split("-")[1]),
                                                  int(list_date_zhuanxie_daili[index_data].split("-")[2])) - datetime.datetime(
                                    int(list_date_zhuanxie_start[index_data].split("-")[0]),
                                    int(list_date_zhuanxie_start[index_data].split("-")[1]),
                                    int(list_date_zhuanxie_start[index_data].split("-")[2]))).days
                            if list_date_zhuanxie_confirm[index_data] == "None":
                                days_zhuanxie_confirm = (
                                datetime.datetime(int(now_year), int(now_month), int(now_days)) - datetime.datetime(
                                    int(list_date_zhuanxie_daili[index_data].split("-")[0]),
                                    int(list_date_zhuanxie_daili[index_data].split("-")[1]),
                                    int(list_date_zhuanxie_daili[index_data].split("-")[2]))).days
                                days_zhuanxie_final = "None"
                            else:
                                days_zhuanxie_confirm = (
                                    datetime.datetime(int(list_date_zhuanxie_confirm[index_data].split("-")[0]),
                                                      int(list_date_zhuanxie_confirm[index_data].split("-")[1]),
                                                      int(list_date_zhuanxie_confirm[index_data].split("-")[2])) - datetime.datetime(
                                        int(list_date_zhuanxie_daili[index_data].split("-")[0]),
                                        int(list_date_zhuanxie_daili[index_data].split("-")[1]),
                                        int(list_date_zhuanxie_daili[index_data].split("-")[2]))).days
                                if list_date_zhuanxie_final[index_data] == "None":
                                    days_zhuanxie_final = (
                                        datetime.datetime(int(now_year), int(now_month), int(now_days)) - datetime.datetime(
                                            int(list_date_zhuanxie_confirm[index_data].split("-")[0]),
                                            int(list_date_zhuanxie_confirm[index_data].split("-")[1]),
                                            int(list_date_zhuanxie_confirm[index_data].split("-")[2]))).days
                                else:
                                    days_zhuanxie_final = (
                                        datetime.datetime(int(list_date_zhuanxie_final[index_data].split("-")[0]),
                                                          int(list_date_zhuanxie_final[index_data].split("-")[1]),
                                                          int(list_date_zhuanxie_final[index_data].split("-")[2])) - datetime.datetime(
                                            int(list_date_zhuanxie_confirm[index_data].split("-")[0]),
                                            int(list_date_zhuanxie_confirm[index_data].split("-")[1]),
                                            int(list_date_zhuanxie_confirm[index_data].split("-")[2]))).days
    list_days_tijiao_jiekouren.append(days_tijiao_jiekouren)
    list_days_tijiao_lingdao.append(days_tijiao_lingdao)
    list_days_tijiao_final.append(days_tijiao_final)
    list_days_zhuanxie_start.append(days_zhuanxie_start)
    list_days_zhuanxie_daili.append(days_zhuanxie_daili)
    list_days_zhuanxie_confirm.append(days_zhuanxie_confirm)
    list_days_zhuanxie_final.append(days_zhuanxie_final)

list_days_tijiao_jiekouren_after_filter = [item for item in list_days_tijiao_jiekouren if item != "None"]
list_days_tijiao_lingdao_after_filter = [item for item in list_days_tijiao_lingdao if item != "None"]
list_days_tijiao_final_after_filter = [item for item in list_days_tijiao_final if item != "None"]
list_days_zhuanxie_start_after_filter = [item for item in list_days_zhuanxie_start if item != "None"]
list_days_zhuanxie_daili_after_filter = [item for item in list_days_zhuanxie_daili if item != "None"]
list_days_zhuanxie_confirm_after_filter = [item for item in list_days_zhuanxie_confirm if item != "None"]
list_days_zhuanxie_final_after_filter = [item for item in list_days_zhuanxie_final if item != "None"]

#分别统计各阶段所耗费时间值。形成一个字典来存储。
dict_days_tijiao_jiekouren = Counter(list_days_tijiao_jiekouren_after_filter)
dict_days_tijiao_lingdao = Counter(list_days_tijiao_lingdao_after_filter)
dict_days_tijiao_final = Counter(list_days_tijiao_final_after_filter)
dict_days_zhuanxie_start = Counter(list_days_zhuanxie_start_after_filter)
dict_days_zhuanxie_daili = Counter(list_days_zhuanxie_daili_after_filter)
dict_days_zhuanxie_confirm = Counter(list_days_zhuanxie_confirm_after_filter)
dict_days_zhuanxie_final = Counter(list_days_zhuanxie_final_after_filter)


days_tijiao_jiekouren_count = []
days_tijiao_lingdao_count = []
days_tijiao_final_count = []
days_zhuanxie_start_count = []
days_zhuanxie_daili_count = []
days_zhuanxie_confirm_count = []
days_zhuanxie_final_count = []

keys_tijiao_jiekouren = sorted(dict_days_tijiao_jiekouren.keys())
keys_tijiao_lingdao = sorted(dict_days_tijiao_lingdao.keys())
keys_tijiao_final = sorted(dict_days_tijiao_final.keys())
keys_zhuanxie_start = sorted(dict_days_zhuanxie_start.keys())
keys_zhuanxie_daili = sorted(dict_days_zhuanxie_daili.keys())
keys_zhuanxie_confirm = sorted(dict_days_zhuanxie_confirm.keys())
keys_zhuanxie_final = sorted(dict_days_zhuanxie_final.keys())

for item in keys_tijiao_jiekouren:
    days_tijiao_jiekouren_count.append(dict_days_tijiao_jiekouren[item])
for item in keys_tijiao_lingdao:
    days_tijiao_lingdao_count.append(dict_days_tijiao_lingdao[item])
for item in keys_tijiao_final:
    days_tijiao_final_count.append(dict_days_tijiao_final[item])
for item in keys_zhuanxie_start:
    days_zhuanxie_start_count.append(dict_days_zhuanxie_start[item])
for item in keys_zhuanxie_daili:
    days_zhuanxie_daili_count.append(dict_days_zhuanxie_daili[item])
for item in keys_zhuanxie_confirm:
    days_zhuanxie_confirm_count.append(dict_days_zhuanxie_confirm[item])
for item in keys_zhuanxie_final:
    days_zhuanxie_final_count.append(dict_days_zhuanxie_final[item])




median_tijiao_jiekouren = numpy.median(numpy.array(list_days_tijiao_jiekouren_after_filter))
median_tijiao_lingdao = numpy.median(numpy.array(list_days_tijiao_lingdao_after_filter))
median_tijiao_final = numpy.median(numpy.array(list_days_tijiao_final_after_filter))
median_zhuanxie_start = numpy.median(numpy.array(list_days_zhuanxie_start_after_filter))
median_zhuanxie_daili = numpy.median(numpy.array(list_days_zhuanxie_daili_after_filter))
median_zhuanxie_confirm = numpy.median(numpy.array(list_days_zhuanxie_confirm_after_filter))
median_zhuanxie_finel = numpy.median(numpy.array(list_days_zhuanxie_final_after_filter))

#画柱状图展现
bar_width = 0.5

#画耗时图像的函数
def plot_image(keys_data, days_data, median_data, filename_image):
    n_groups = len(keys_data)
    figure = pyplot.figure("%s" % filename_image)
    sub_figure = figure.add_subplot()
    index = numpy.arange(n_groups)
    pyplot.bar(left=index, height=days_data, width=bar_width, color='r', label='Days Used By %s' %filename_image)
    pyplot.xlabel('Days Used')
    pyplot.ylabel('Count')
    pyplot.xticks(index, keys_data)
    pyplot.axvline(median_data)
    pyplot.legend()
    pyplot.tight_layout()
    pyplot.show()
    #pyplot.savefig("Days_Used_By_%s.png" % filename_image)

#plot_image(keys_tijiao_jiekouren, days_tijiao_jiekouren_count, median_tijiao_jiekouren, "Jiekouren")
#plot_image(keys_tijiao_lingdao, days_tijiao_lingdao_count, median_tijiao_lingdao, "LingDao")
#plot_image(keys_tijiao_final, days_tijiao_final_count, median_tijiao_final, "ZhuanliEngineer")
#plot_image(keys_zhuanxie_start, days_zhuanxie_start_count, median_zhuanxie_start, "StartToWrite")
plot_image(keys_zhuanxie_daili, days_zhuanxie_daili_count, median_zhuanxie_daili, "DaiLi")
#plot_image(keys_zhuanxie_confirm, days_zhuanxie_confirm_count, median_zhuanxie_confirm, "CreatorConfirm")
#plot_image(keys_zhuanxie_final, days_zhuanxie_final_count, median_zhuanxie_finel, "ZhuanxieFinal")

