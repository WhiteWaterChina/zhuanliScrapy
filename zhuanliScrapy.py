#!/usr/bin/env python
# -*- coding:cp936 -*-
#author:yanshuo@inspur.com
import time
import os
import sys
import re
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import tkMessageBox
import xlsxwriter
import datetime
import Tkinter


def getdatafromweb():
    data_status_list = []
    data_sn_list = []
    data_filename_list = []
    data_creator_list = []
    data_created_date_list = []
    data_current_nodename_list = []
    department_name_list = []
    type_invention_list = []
    shouli_sn_list = []
    username = 'yanshuo@inspur.com'
    password = 'patsnapinspur'
    driverpath = os.path.join(os.path.abspath(os.path.curdir), "phantomjs.exe")
    browser = webdriver.PhantomJS(driverpath)
    url = "http://10.110.6.34/users/login"
    browser.get(url)
    browser.find_element_by_id("UserEmail").send_keys(username)
    browser.find_element_by_id("EmailPassword").send_keys(password)
    browser.find_element_by_css_selector("button.new-login").click()
    time.sleep(5)
    browser.find_element_by_css_selector("#header > ul > li:nth-child(2)")
    ActionChains(browser).move_to_element(browser.find_element_by_css_selector("#header > ul > li:nth-child(2)")).perform()
    browser.find_element_by_css_selector("#header > ul > li:nth-child(2) > div > div > ul > li:nth-child(2) > a").click()
    time.sleep(3)
    except_list = ['驳回'.decode('gbk'), '退回发起人'.decode('gbk'), '撤销'.decode('gbk')]
    while True:
        current_table_line = browser.find_elements_by_css_selector("#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr")
        length_table = len(current_table_line) + 1
        for line_number in range(1, length_table):
            data_status = browser.find_element_by_css_selector("#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.status > span" % line_number).text
            data_sn_filename_link = browser.find_element_by_css_selector("#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.subject > a " % line_number)
            data_sn_filename = data_sn_filename_link.text
            data_sn = data_sn_filename.split('/')[0].strip()
            if data_status not in except_list and data_sn not in data_sn_list:
                data_current_nodename = browser.find_element_by_css_selector("#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.node_name" % line_number).text.strip()
#                data_created_by = browser.find_element_by_css_selector("#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.created_by" % line_number).text.strip()
                data_created_at_temp = browser.find_element_by_css_selector("#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.created_at" % line_number).text.strip()
                data_created_at = data_created_at_temp
                data_sn_list.append(data_sn)
                data_current_nodename_list.append(data_current_nodename)
#                data_creator_list.append(data_created_by)
                data_created_date_list.append(data_created_at)
                data_sn_filename_link.click()
                time.sleep(3)
                handles = browser.window_handles
                browser.switch_to.window(handles[1])
                WebDriverWait(browser, 100).until(ec.presence_of_element_located((By.CSS_SELECTOR, '#main > div.major > div.major-section.clearfix > div.content-wrapper.clearfix.layout-detail-main > div.basic-info > div.major-left > div > table > tbody > tr:nth-child(10) > th')))
                try:
                    departmane_name = browser.find_element_by_css_selector("#main > div.major > div.major-section.clearfix > div.content-wrapper.clearfix.layout-detail-main > div.basic-info > div.major-left > div > table > tbody > tr:nth-child(10) > td > a:nth-child(4)").text.strip()
                except selenium.common.exceptions.NoSuchElementException:
                    departmane_name = 'None'

                type_invention = browser.find_element_by_css_selector('#main > div.major > div.major-section.clearfix > div.content-wrapper.clearfix.layout-detail-main > div.basic-info > div.major-left > div > table > tbody > tr:nth-child(6) > td').text.strip()
                data_status_display = browser.find_element_by_css_selector("#main > div.major > div.major-section.clearfix > div.major-header > div.major-title > span").text.strip()
                data_created_by = browser.find_element_by_css_selector("#main > div.major > div.major-section.clearfix > div.content-wrapper.clearfix.layout-detail-main > div.basic-info > div.major-left > div > table > tbody > tr:nth-child(20) > td").text.split(" ")[0].strip()
                data_creator_list.append(data_created_by)
                if data_status_display == '申请专利'.decode('gbk'):
                    shouli_sn = browser.find_element_by_css_selector("#patents-related > div > span.table-content > table > tbody > tr > td:nth-child(3)").text.strip()
                    shouli_sn_list.append(shouli_sn)
                else:
                    shouli_sn_list.append('None')
                data_filaname = browser.find_element_by_css_selector("#main > div.major > div.major-section.clearfix > div.content-wrapper.clearfix.layout-detail-main > div.basic-info > div.major-left > div > table > tbody > tr:nth-child(2) > td").text.strip()
                data_filename_list.append(data_filaname)
                department_name_list.append(departmane_name)
                type_invention_list.append(type_invention)
                data_status_list.append(data_status_display)
                browser.close()
                browser.switch_to.window(handles[0])
        current_page_number = browser.find_element_by_css_selector("#table_page > div > span").text.strip()
        print "处理完成第%s页".decode('gbk') %current_page_number
        try:
            total_bottom_div = len(browser.find_elements_by_css_selector("#table_page > div > a"))
            next_page = browser.find_element_by_css_selector("#table_page > div > a:nth-child(%d)" % total_bottom_div)
            if next_page.text != "下一页".decode('gbk'):
                browser.quit()
                break
            else:
                next_page.click()
                time.sleep(3)
        except selenium.common.exceptions.NoSuchElementException:
            browser.quit()
            break

    return data_status_list, data_sn_list, data_filename_list, department_name_list, type_invention_list, data_current_nodename_list, data_creator_list, data_created_date_list, shouli_sn_list


def write_excel(data_status_list, data_sn_list, data_filename_list, department_name_list, type_invention_list, data_current_nodename_list, data_creator_list, data_created_date_list, shouli_sn_list):
    title_sheet = ['当前状态'.decode('gbk'), '提案编号'.decode('gbk'), '提案名称'.decode('gbk'), '处别'.decode('gbk'), '发明类型'.decode('gbk'), '当前处理节点'.decode('gbk'), '创建者'.decode('gbk'), '创建时间'.decode('gbk'), '受理申请编号'.decode('gbk')]
    timestamp = time.strftime('%Y%m%d', time.localtime())
    workbook_display = xlsxwriter.Workbook('测试验证部专利总览-%s.xlsx'.decode('gbk') % timestamp)
    sheet = workbook_display.add_worksheet('2017财年测试验证部专利统计'.decode('gbk'))
    formatOne = workbook_display.add_format()
    formatOne.set_border(1)
    formatTwo = workbook_display.add_format()
    formatTwo.set_border(1)
    formattitle = workbook_display.add_format()
    formattitle.set_border(1)
    formattitle.set_align('center')
    formattitle.set_bg_color("yellow")
    formattitle.set_bold(True)
    sheet.set_column('H:I', 22)
    sheet.set_column('B:B', 14)
    sheet.set_column('C:C', 58)
    sheet.merge_range(0, 0, 0, 8, "测试验证部2017财年专利总览".decode('gbk'), formattitle)
    for index_title, item_title in enumerate(title_sheet):
        sheet.write(1, index_title, item_title, formatOne)
        for index_data, item_data in enumerate(data_sn_list):
            sheet.write(2 + index_data, 0, data_status_list[index_data], formatOne)
            sheet.write(2 + index_data, 1, data_sn_list[index_data], formatOne)
            sheet.write(2 + index_data, 2, data_filename_list[index_data], formatOne)
            sheet.write(2 + index_data, 3, department_name_list[index_data], formatOne)
            sheet.write(2 + index_data, 4, type_invention_list[index_data], formatOne)
            sheet.write(2 + index_data, 5, data_current_nodename_list[index_data], formatOne)
            sheet.write(2 + index_data, 6, data_creator_list[index_data], formatOne)
            sheet.write_datetime(2 + index_data, 7, datetime.datetime.strptime(data_created_date_list[index_data], '%Y/%m/%d %H:%M:%S'), workbook_display.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss', 'border': 1}))
            sheet.write(2 + index_data, 8, shouli_sn_list[index_data], formatOne)
    workbook_display.close()
print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
a, b, c, d, e, f, g, h, i = getdatafromweb()
write_excel(a, b, c, d, e, f, g, h, i)
print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
