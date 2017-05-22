#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import wx
import wx.xrc
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


class FrameZhuanli(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"专利信息扒取系统", pos=wx.DefaultPosition, size=wx.Size(418, 297),
                          style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)
        self.SetFont(wx.Font(9, 70, 90, 90, False, "宋体"))
        self.SetForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOWFRAME))

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        bSizer2 = wx.BoxSizer(wx.HORIZONTAL)

        self.text_department = wx.StaticText(self, wx.ID_ANY, u"请输入部门名称", wx.DefaultPosition, wx.Size(150, 20),
                                             wx.ALIGN_CENTRE)
        self.text_department.Wrap(-1)
        self.text_department.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_department.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer2.Add(self.text_department, 0, wx.ALL, 5)

        self.input_department = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.input_department, 1, wx.ALL, 5)

        bSizer1.Add(bSizer2, 0, wx.EXPAND, 5)

        bSizer3 = wx.BoxSizer(wx.HORIZONTAL)

        self.text_username = wx.StaticText(self, wx.ID_ANY, u"请输入用户名", wx.DefaultPosition, wx.Size(150, 20),
                                           wx.ALIGN_CENTRE)
        self.text_username.Wrap(-1)
        self.text_username.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_username.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer3.Add(self.text_username, 0, wx.ALL, 5)

        self.input_username = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer3.Add(self.input_username, 1, wx.ALL, 5)

        bSizer1.Add(bSizer3, 0, wx.EXPAND, 5)

        bSizer4 = wx.BoxSizer(wx.HORIZONTAL)

        self.text_password = wx.StaticText(self, wx.ID_ANY, u"请输入密码", wx.DefaultPosition, wx.Size(150, 20),
                                           wx.ALIGN_CENTRE)
        self.text_password.Wrap(-1)
        self.text_password.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_password.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer4.Add(self.text_password, 0, wx.ALL, 5)

        self.input_password = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          wx.TE_PASSWORD)
        bSizer4.Add(self.input_password, 1, wx.ALL, 5)

        bSizer1.Add(bSizer4, 0, wx.EXPAND, 5)

        bSizer8 = wx.BoxSizer(wx.HORIZONTAL)

        bSizer9 = wx.BoxSizer(wx.VERTICAL)

        bSizer8.Add(bSizer9, 1, wx.EXPAND, 5)

        self.test_info = wx.StaticText(self, wx.ID_ANY, u"请选择排除在外的状态", wx.DefaultPosition, wx.DefaultSize, 0)
        self.test_info.Wrap(-1)
        self.test_info.SetForegroundColour(wx.Colour(255, 255, 0))
        self.test_info.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer8.Add(self.test_info, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10 = wx.BoxSizer(wx.HORIZONTAL)

        self.checkbox_1 = wx.CheckBox(self, wx.ID_ANY, u"撤销", wx.DefaultPosition, wx.Size(60, -1), 0)
        self.checkbox_1.SetValue(True)
        self.checkbox_1.SetFont(wx.Font(9, 70, 90, 90, False, "宋体"))
        self.checkbox_1.SetForegroundColour(wx.Colour(0, 255, 64))

        bSizer10.Add(self.checkbox_1, 0, wx.ALL, 5)

        self.checkbox_2 = wx.CheckBox(self, wx.ID_ANY, u"退回发起人", wx.DefaultPosition, wx.Size(80, -1), 0)
        self.checkbox_2.SetValue(True)
        self.checkbox_2.SetForegroundColour(wx.Colour(0, 255, 64))

        bSizer10.Add(self.checkbox_2, 0, wx.ALL, 5)

        self.checkbox_3 = wx.CheckBox(self, wx.ID_ANY, u"驳回", wx.DefaultPosition, wx.Size(80, -1), 0)
        self.checkbox_3.SetValue(True)
        self.checkbox_3.SetForegroundColour(wx.Colour(0, 255, 64))

        bSizer10.Add(self.checkbox_3, 0, wx.ALL, 5)

        bSizer8.Add(bSizer10, 1, wx.EXPAND, 5)

        bSizer1.Add(bSizer8, 0, wx.EXPAND, 5)

        bSizer5 = wx.BoxSizer(wx.HORIZONTAL)

        self.button_go = btn = wx.Button(self, wx.ID_ANY, u"GO", wx.DefaultPosition, wx.Size(-1, 35), 0)
        bSizer5.Add(self.button_go, 0, wx.ALL, 5)

        self.button_exit = wx.Button(self, wx.ID_ANY, u"退出", wx.DefaultPosition, wx.Size(-1, 35), 0)
        bSizer5.Add(self.button_exit, 0, wx.ALL, 5)

        bSizer1.Add(bSizer5, 0, wx.ALIGN_CENTER | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer91 = wx.BoxSizer(wx.VERTICAL)

        self.output_info = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_MULTILINE)
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
        data_status_list = []
        data_sn_list = []
        data_filename_list = []
        data_creator_list = []
        data_created_date_list = []
        data_current_nodename_list = []
        department_name_list = []
        type_invention_list = []
        shouli_sn_list = []
        driverpath = os.path.join(os.path.abspath(os.path.curdir), "phantomjs.exe")
        browser = webdriver.PhantomJS(driverpath)
        url = "http://10.110.6.34/users/login"
        browser.get(url)
        browser.find_element_by_id("UserEmail").send_keys(username)
        browser.find_element_by_id("EmailPassword").send_keys(password)
        browser.find_element_by_css_selector("button.new-login").click()
        time.sleep(5)
        browser.find_element_by_css_selector("#header > ul > li:nth-child(2)")
        ActionChains(browser).move_to_element(
            browser.find_element_by_css_selector("#header > ul > li:nth-child(2)")).perform()
        browser.find_element_by_css_selector(
            "#header > ul > li:nth-child(2) > div > div > ul > li:nth-child(2) > a").click()
        time.sleep(3)
        except_list = []
        if self.checkbox_1.GetValue():
            except_list.append('驳回'.decode('gbk'))
        if self.checkbox_2.GetValue():
            except_list.append('退回发起人'.decode('gbk'))
        if self.checkbox_3.GetValue():
            except_list.append('撤销'.decode('gbk'))
        while True:
            current_table_line = browser.find_elements_by_css_selector(
                "#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr")
            length_table = len(current_table_line) + 1
            for line_number in range(1, length_table):
                data_status = browser.find_element_by_css_selector(
                    "#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.status > span" % line_number).text
                data_sn_filename_link = browser.find_element_by_css_selector(
                    "#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.subject > a " % line_number)
                data_sn_filename = data_sn_filename_link.text
                data_sn = data_sn_filename.split('/')[0].strip()
                if data_status not in except_list and data_sn not in data_sn_list:
                    data_current_nodename = browser.find_element_by_css_selector(
                        "#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.node_name" % line_number).text.strip()
                    data_created_at_temp = browser.find_element_by_css_selector(
                        "#list-result > div.template-list-condition > div.list-mail-con > table > tbody > tr:nth-child(%d) > td.cos.created_at" % line_number).text.strip()
                    data_created_at = data_created_at_temp
                    data_sn_list.append(data_sn)
                    data_current_nodename_list.append(data_current_nodename)
                    data_created_date_list.append(data_created_at)
                    data_sn_filename_link.click()
                    time.sleep(3)
                    handles = browser.window_handles
                    browser.switch_to.window(handles[1])
                    WebDriverWait(browser, 100).until(ec.presence_of_element_located((By.CSS_SELECTOR,
                                                                                      '#main > div.major > div.major-section.clearfix > div.content-wrapper.clearfix.layout-detail-main > div.basic-info > div.major-left > div > table > tbody > tr:nth-child(10) > th')))
                    department_temp = browser.find_elements_by_css_selector(
                        "#main > div.major > div.major-section.clearfix > div.content-wrapper.clearfix.layout-detail-main > div.basic-info > div.major-left > div > table > tbody > tr:nth-child(10) > td > a")
                    #length_department = len(department_temp)
                    #department_name = browser.find_element_by_css_selector(
                        #"#main > div.major > div.major-section.clearfix > div.content-wrapper.clearfix.layout-detail-main > div.basic-info > div.major-left > div > table > tbody > tr:nth-child(10) > td > a:nth-child(%d)" % length_department).text.strip()
                    department_name_temp_list = []
                    for item_department in department_temp:
                        department_name_temp_list.append(item_department.text.strip())
                    department_name = "".join(department_name_temp_list)
                    type_invention = browser.find_element_by_css_selector(
                        '#main > div.major > div.major-section.clearfix > div.content-wrapper.clearfix.layout-detail-main > div.basic-info > div.major-left > div > table > tbody > tr:nth-child(6) > td').text.strip()
                    data_status_display = browser.find_element_by_css_selector(
                        "#main > div.major > div.major-section.clearfix > div.major-header > div.major-title > span").text.strip()
                    data_created_by = browser.find_element_by_css_selector(
                        "#main > div.major > div.major-section.clearfix > div.content-wrapper.clearfix.layout-detail-main > div.basic-info > div.major-left > div > table > tbody > tr:nth-child(20) > td").text.split(
                        " ")[0].strip()
                    data_creator_list.append(data_created_by)
                    if data_status_display == '申请专利'.decode('gbk'):
                        shouli_sn = browser.find_element_by_css_selector(
                            "#patents-related > div > span.table-content > table > tbody > tr > td:nth-child(3)").text.strip()
                        shouli_sn_list.append(shouli_sn)
                    else:
                        shouli_sn_list.append('None')
                    data_filaname = browser.find_element_by_css_selector(
                        "#main > div.major > div.major-section.clearfix > div.content-wrapper.clearfix.layout-detail-main > div.basic-info > div.major-left > div > table > tbody > tr:nth-child(2) > td").text.strip()
                    data_filename_list.append(data_filaname)
                    department_name_list.append(department_name)
                    type_invention_list.append(type_invention)
                    data_status_list.append(data_status_display)
                    browser.close()
                    browser.switch_to.window(handles[0])
            current_page_number = int(browser.find_element_by_css_selector("#table_page > div > span").text.strip())
            self.updatedisplay(current_page_number)
            try:
                total_bottom_div = len(browser.find_elements_by_css_selector("#table_page > div > a"))
                next_page = browser.find_element_by_css_selector(
                    "#table_page > div > a:nth-child(%d)" % total_bottom_div)
                if next_page.text != "下一页".decode('gbk'):
                    browser.quit()
                    break
                else:
                    next_page.click()
                    time.sleep(3)
            except selenium.common.exceptions.NoSuchElementException:
                browser.quit()
                break

        title_sheet = ['当前状态'.decode('gbk'), '提案编号'.decode('gbk'), '提案名称'.decode('gbk'), '处别'.decode('gbk'),
                       '发明类型'.decode('gbk'), '当前处理节点'.decode('gbk'), '创建者'.decode('gbk'), '创建时间'.decode('gbk'),
                       '受理申请编号'.decode('gbk')]
        timestamp = time.strftime('%Y%m%d', time.localtime())
        workbook_display = xlsxwriter.Workbook('%s专利总览-%s.xlsx'.decode('gbk') % (department_write, timestamp))
        sheet = workbook_display.add_worksheet('2017财年%s专利统计'.decode('gbk') % department_write)
        formatone = workbook_display.add_format()
        formatone.set_border(1)
        formattwo = workbook_display.add_format()
        formattwo.set_border(1)
        formattitle = workbook_display.add_format()
        formattitle.set_border(1)
        formattitle.set_align('center')
        formattitle.set_bg_color("yellow")
        formattitle.set_bold(True)
        sheet.set_column('H:I', 22)
        sheet.set_column('B:B', 14)
        sheet.set_column('C:C', 42)
        sheet.set_column('D:D', 34)
        sheet.merge_range(0, 0, 0, 8, "%s2017财年专利总览".decode('gbk') % department_write, formattitle)
        for index_title, item_title in enumerate(title_sheet):
            sheet.write(1, index_title, item_title, formatone)
            for index_data, item_data in enumerate(data_sn_list):
                sheet.write(2 + index_data, 0, data_status_list[index_data], formatone)
                sheet.write(2 + index_data, 1, data_sn_list[index_data], formatone)
                sheet.write(2 + index_data, 2, data_filename_list[index_data], formatone)
                sheet.write(2 + index_data, 3, department_name_list[index_data], formatone)
                sheet.write(2 + index_data, 4, type_invention_list[index_data], formatone)
                sheet.write(2 + index_data, 5, data_current_nodename_list[index_data], formatone)
                sheet.write(2 + index_data, 6, data_creator_list[index_data], formatone)
                sheet.write_datetime(2 + index_data, 7, datetime.datetime.strptime(data_created_date_list[index_data],
                                                                                   '%Y/%m/%d %H:%M:%S'),
                                     workbook_display.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss', 'border': 1}))
                sheet.write(2 + index_data, 8, shouli_sn_list[index_data], formatone)
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
            self.output_info.AppendText("完成第%s页" % t)
        elif t == "Finished":
            self.button_go.Enable()
        else:
            self.output_info.AppendText("%s" % t)
        self.output_info.AppendText(os.linesep)


if __name__ == '__main__':
    app = wx.App()
    frame = FrameZhuanli(None)
    frame.Show()
    app.MainLoop()