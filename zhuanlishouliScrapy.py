#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import wx
import time
import os
from threading import Thread
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import xlsxwriter
import datetime
import re


class FrameZhuanli(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"专利信息扒取系统", pos=wx.DefaultPosition, size=wx.Size(393, 411),
                          style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetFont(wx.Font(wx.NORMAL_FONT.GetPointSize(), 70, 90, 90, False, wx.EmptyString))
        self.SetForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOWFRAME))

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        bSizer101 = wx.BoxSizer(wx.VERTICAL)

        self.m_panel1 = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer12 = wx.BoxSizer(wx.VERTICAL)

        bSizer2 = wx.BoxSizer(wx.HORIZONTAL)

        self.text_department = wx.StaticText(self.m_panel1, wx.ID_ANY, u"请输入部门名称", wx.DefaultPosition, wx.Size(150, 20),
                                             wx.ALIGN_CENTRE)
        self.text_department.Wrap(-1)
        self.text_department.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_department.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer2.Add(self.text_department, 0, wx.ALL, 5)

        self.input_department = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        bSizer2.Add(self.input_department, 1, wx.ALL, 5)

        bSizer12.Add(bSizer2, 1, wx.EXPAND, 5)

        bSizer3 = wx.BoxSizer(wx.HORIZONTAL)

        self.text_username = wx.StaticText(self.m_panel1, wx.ID_ANY, u"请输入用户名", wx.DefaultPosition, wx.Size(150, 20),
                                           wx.ALIGN_CENTRE)
        self.text_username.Wrap(-1)
        self.text_username.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_username.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer3.Add(self.text_username, 0, wx.ALL, 5)

        self.input_username = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          0)
        bSizer3.Add(self.input_username, 1, wx.ALL, 5)

        bSizer12.Add(bSizer3, 0, wx.EXPAND, 5)

        bSizer4 = wx.BoxSizer(wx.HORIZONTAL)

        self.text_password = wx.StaticText(self.m_panel1, wx.ID_ANY, u"请输入密码", wx.DefaultPosition, wx.Size(150, 20),
                                           wx.ALIGN_CENTRE)
        self.text_password.Wrap(-1)
        self.text_password.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_password.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer4.Add(self.text_password, 0, wx.ALL, 5)

        self.input_password = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          wx.TE_PASSWORD)
        bSizer4.Add(self.input_password, 1, wx.ALL, 5)

        bSizer12.Add(bSizer4, 0, wx.EXPAND, 5)

        bSizer14 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText5 = wx.StaticText(self.m_panel1, wx.ID_ANY,
                                           u"请在如下输入需要抓取专利信息的开始和结束日期！\n日期格式20170731.个位数的日期一定要补全0！", wx.DefaultPosition,
                                           wx.DefaultSize, 0)
        self.m_staticText5.Wrap(-1)
        self.m_staticText5.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText5.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer14.Add(self.m_staticText5, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer12.Add(bSizer14, 0, wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer16 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText6 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"开始日期:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText6.Wrap(-1)
        self.m_staticText6.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText6.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer16.Add(self.m_staticText6, 0, wx.ALL, 5)

        self.text_startdate = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          0)
        bSizer16.Add(self.text_startdate, 0, wx.ALL, 5)

        self.m_staticText7 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"结束日期:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText7.Wrap(-1)
        self.m_staticText7.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText7.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer16.Add(self.m_staticText7, 0, wx.ALL, 5)

        self.text_enddate = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer16.Add(self.text_enddate, 0, wx.ALL, 5)

        bSizer12.Add(bSizer16, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_panel1.SetSizer(bSizer12)
        self.m_panel1.Layout()
        bSizer12.Fit(self.m_panel1)
        bSizer101.Add(self.m_panel1, 1, wx.EXPAND | wx.ALL, 5)

        bSizer1.Add(bSizer101, 1, wx.EXPAND, 5)

        bSizer8 = wx.BoxSizer(wx.HORIZONTAL)

        bSizer9 = wx.BoxSizer(wx.VERTICAL)

        bSizer8.Add(bSizer9, 1, wx.EXPAND, 5)

        self.test_info = wx.StaticText(self, wx.ID_ANY, u"请选择排除在外的状态", wx.DefaultPosition, wx.DefaultSize, 0)
        self.test_info.Wrap(-1)
        self.test_info.SetForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_INFOTEXT))
        self.test_info.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer8.Add(self.test_info, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_panel2 = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer15 = wx.BoxSizer(wx.HORIZONTAL)

        self.checkbox_1 = wx.CheckBox(self.m_panel2, wx.ID_ANY, u"撤销", wx.DefaultPosition, wx.Size(-1, -1), 0)
        self.checkbox_1.SetValue(True)
        self.checkbox_1.SetFont(wx.Font(wx.NORMAL_FONT.GetPointSize(), 70, 90, 90, False, wx.EmptyString))
        self.checkbox_1.SetForegroundColour(wx.Colour(0, 255, 64))

        bSizer15.Add(self.checkbox_1, 0, wx.ALL, 5)

        self.checkbox_2 = wx.CheckBox(self.m_panel2, wx.ID_ANY, u"退回发起人", wx.DefaultPosition, wx.Size(-1, -1), 0)
        self.checkbox_2.SetValue(True)
        self.checkbox_2.SetForegroundColour(wx.Colour(0, 255, 64))

        bSizer15.Add(self.checkbox_2, 0, wx.ALL, 5)

        self.checkbox_3 = wx.CheckBox(self.m_panel2, wx.ID_ANY, u"驳回", wx.DefaultPosition, wx.Size(-1, -1), 0)
        self.checkbox_3.SetValue(True)
        self.checkbox_3.SetForegroundColour(wx.Colour(0, 255, 64))

        bSizer15.Add(self.checkbox_3, 0, wx.ALL, 5)

        self.m_panel2.SetSizer(bSizer15)
        self.m_panel2.Layout()
        bSizer15.Fit(self.m_panel2)
        bSizer10.Add(self.m_panel2, 1, wx.EXPAND | wx.ALL, 5)

        bSizer8.Add(bSizer10, 0, wx.EXPAND, 5)

        bSizer1.Add(bSizer8, 0, wx.EXPAND, 5)

        bSizer5 = wx.BoxSizer(wx.HORIZONTAL)

        self.button_go = wx.Button(self, wx.ID_ANY, u"GO", wx.DefaultPosition, wx.Size(-1, 35), 0)
        bSizer5.Add(self.button_go, 0, wx.ALL, 5)

        self.button_exit = wx.Button(self, wx.ID_ANY, u"退出", wx.DefaultPosition, wx.Size(-1, 35), 0)
        bSizer5.Add(self.button_exit, 0, wx.ALL, 5)

        bSizer1.Add(bSizer5, 0, wx.ALIGN_CENTER | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer91 = wx.BoxSizer(wx.VERTICAL)

        self.output_info = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                       wx.TE_MULTILINE | wx.TE_READONLY)
        bSizer91.Add(self.output_info, 1, wx.ALL | wx.EXPAND, 5)

        bSizer1.Add(bSizer91, 1, wx.EXPAND, 5)

        self.SetSizer(bSizer1)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.button_go.Bind(wx.EVT_BUTTON, self.onbutton)
        self.button_exit.Bind(wx.EVT_BUTTON, self.close)

        self._thread = Thread(target=self.run, args=())
        self._thread.daemon = True

    def close(self, event):
        self.Close()

    def run(self):
        self.updatedisplay("开始抓取".decode('gbk'))
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        department_write = self.input_department.GetValue()
        username = self.input_username.GetValue()
        password = self.input_password.GetValue()
        startdate = self.text_startdate.GetValue().strip()
        if len(startdate) == 0:
            startdate = "20170301"
        enddate = self.text_enddate.GetValue().strip()
        if len(enddate) == 0:
            enddate = int(time.strftime('%Y%m%d', time.localtime(time.time()))) + 1
        data_management_sn_list = []
        data_sn_list = []
        data_filename_original_list = []
        data_filename_final_list = []
        data_creator_list = []
        data_created_date_list = []
        department_name_list = []
        type_invention_list = []

        driverpath = os.path.join(os.path.abspath(os.path.curdir), "chromedriver.exe")
        browser = webdriver.Chrome(driverpath)
        #        driverpath = os.path.join(os.path.abspath(os.path.curdir), "phantomjs.exe")
        #        browser = webdriver.PhantomJS(driverpath)
        url = "http://10.110.6.34/users/login"
        browser.get(url)
        browser.find_element_by_id("UserEmail").send_keys(username)
        browser.find_element_by_id("EmailPassword").send_keys(password)
        browser.find_element_by_css_selector("button.new-login").click()
        time.sleep(5)
        browser.find_element_by_css_selector("#header > ul > li:nth-child(4)")
        ActionChains(browser).move_to_element(
            browser.find_element_by_css_selector("#header > ul > li:nth-child(4)")).perform()
        browser.find_element_by_css_selector(
            "#header > ul > li:nth-child(4) > div > div > ul > li:nth-child(1) > a").click()
        time.sleep(3)

        while True:
            current_table_line = browser.find_elements_by_css_selector(
                "#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr")
            length_table = len(current_table_line) + 1
            for line_number in range(1, length_table):
                # 获取首页面显示的信息
                # 获取管理SN
                data_management_sn = browser.find_element_by_css_selector(
                    "#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.patent_management_sn" % line_number).text.strip()
                print data_management_sn
                # 获取专利类型，发明还是实用
                data_type = browser.find_element_by_css_selector(
                    "#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.patent_patent_name > span" % line_number).text.strip()
                # 获取链接
                data_name_link = browser.find_element_by_css_selector(
                    "#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.patent_patent_name > a" % line_number)
                # 获取经过代理撰写后的专利名称
                data_name_final = data_name_link.text.strip()
                # 获取受理编号
                data_sn = browser.find_element_by_css_selector(
                    "#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.preliminarybase_application_number" % line_number).text.strip()
                # 获取受理时间
                data_created_at_temp = browser.find_element_by_css_selector(
                    "#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.preliminarybase_filed_date" % line_number).text.strip()
                list_data_created_at_limit = data_created_at_temp.split(" ")[0].split("/")
                data_created_at_limit = "".join(list_data_created_at_limit)

                # 排除掉不在时间范围内的和management_sn重复的
                if data_sn not in data_sn_list and int(startdate) < int(data_created_at_limit) < int(enddate):
                    data_name_link.click()
                    time.sleep(3)
                    handles = browser.window_handles
                    browser.switch_to.window(handles[1])
                    try:
                        WebDriverWait(browser, 30).until(ec.presence_of_element_located(
                            (By.CSS_SELECTOR, '#patentDetail > table > tbody > tr:nth-child(30) > td > div > div')))
                    except selenium.common.exceptions.TimeoutException:
                        browser.close()
                        browser.switch_to.window(handles[0])
                        continue
                    # 获取处级别
                    department_temp = browser.find_elements_by_css_selector(
                        "#patentDetail > table > tbody > tr:nth-child(31) > td > div > div > a")
                    department_name_temp_list = []
                    for item_department in department_temp:
                        department_name_temp_list.append(item_department.text.strip())
                    department_name = "".join(department_name_temp_list)
                    # 获取发明原始名称
                    data_name_original = browser.find_element_by_css_selector(
                        "#patentDetail > table > tbody > tr:nth-child(5) > td > div > div").text.strip()
                    # 获取发明人
                    data_created_by_temp = browser.find_element_by_css_selector(
                        "#patentDetail > table > tbody > tr:nth-child(30) > td > div > div").text.split(" ")[0].strip()
                    data_created_by = re.search(r"\D*", data_created_by_temp).group()
                    # 将数据写入list
                    data_management_sn_list.append(data_management_sn)
                    data_filename_original_list.append(data_name_original)
                    data_filename_final_list.append(data_name_final)
                    type_invention_list.append(data_type)
                    data_creator_list.append(data_created_by)
                    data_sn_list.append(data_sn)
                    data_created_date_list.append(data_created_at_temp)
                    department_name_list.append(department_name)
                    browser.close()
                    browser.switch_to.window(handles[0])
            current_page_number = int(browser.find_element_by_css_selector("#table_page > div > span").text.strip())
            self.updatedisplay(current_page_number)
            try:
                total_bottom_div = len(browser.find_elements_by_css_selector("#table_page > div > a"))
                next_page = browser.find_element_by_css_selector(
                    "#table_page > div > a:nth-child(%d)" % total_bottom_div)
                if next_page.text.strip() != "下一页".decode('gbk'):
                    browser.quit()
                    break
                else:
                    next_page.click()
                    time.sleep(3)
                    WebDriverWait(browser, 100).until(ec.presence_of_element_located((By.CSS_SELECTOR, '#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(1) > td.cos.patent_management_sn')))
            except selenium.common.exceptions.NoSuchElementException:
                browser.quit()
                break

        title_sheet = ['管理编号'.decode('gbk'), '专利类型'.decode('gbk'), '原专利名称'.decode('gbk'), '代理提交专利名称'.decode('gbk'), '发明人'.decode('gbk'), '申请号'.decode('gbk'), '申请日期'.decode('gbk'), '部门'.decode('gbk')]
        timestamp = time.strftime('%Y%m%d', time.localtime())
        workbook_display = xlsxwriter.Workbook('%s专利受理总览-%s.xlsx'.decode('gbk') % (department_write, timestamp))
        sheet = workbook_display.add_worksheet('2017财年%s受理专利统计'.decode('gbk') % department_write)
        formatone = workbook_display.add_format()
        formatone.set_border(1)
        formattwo = workbook_display.add_format()
        formattwo.set_border(1)
        formattitle = workbook_display.add_format()
        formattitle.set_border(1)
        formattitle.set_align('center')
        formattitle.set_bg_color("yellow")
        formattitle.set_bold(True)
        sheet.set_column('A:A', 17)
        sheet.set_column('C:D', 42)
        sheet.set_column('G:G', 13)
        sheet.set_column('F:F', 18)
        sheet.merge_range(0, 0, 0, 8, "%s2017财年受理专利总览".decode('gbk') % department_write, formattitle)
        for index_title, item_title in enumerate(title_sheet):
            sheet.write(1, index_title, item_title, formatone)
            for index_data, item_data in enumerate(data_sn_list):
                sheet.write(2 + index_data, 0, data_management_sn_list[index_data], formatone)
                sheet.write(2 + index_data, 1, type_invention_list[index_data], formatone)
                sheet.write(2 + index_data, 2, data_filename_original_list[index_data], formatone)
                sheet.write(2 + index_data, 3, data_filename_final_list[index_data], formatone)
                sheet.write(2 + index_data, 4, data_creator_list[index_data], formatone)
                sheet.write(2 + index_data, 5, data_sn_list[index_data], formatone)
                sheet.write_datetime(2 + index_data, 6, datetime.datetime.strptime(data_created_date_list[index_data],
                                                                                   '%Y/%m/%d'),
                                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
                sheet.write(2 + index_data, 7, department_name_list[index_data], formatone)
        workbook_display.close()
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        self.updatedisplay("抓取结束,请点击退出按钮退出程序".decode('gbk'))
        time.sleep(1)
        self.updatedisplay("Finished")

    def onbutton(self, event):
        self._thread.start()
        self.started = True
        self.button_go = event.GetEventObject()
        self.button_go.Disable()

    def updatedisplay(self, msg):
        t = msg
        if isinstance(t, int):
            self.output_info.AppendText("完成第%s页".decode('gbk') % t)
        elif t == "Finished":
            self.button_go.Enable()
        else:
            self.output_info.AppendText("%s".decode('gbk') % t)
        self.output_info.AppendText(os.linesep)


if __name__ == '__main__':
    app = wx.App()
    frame = FrameZhuanli(None)
    frame.Show()
    app.MainLoop()
