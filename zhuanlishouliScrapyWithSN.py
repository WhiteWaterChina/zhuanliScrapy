#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import wx
import time
import os
from threading import Thread
import xlsxwriter
import datetime
import re
import urllib2
from bs4 import BeautifulSoup
import requests


class FrameZhuanli(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"专利受理信息扒取系统", pos=wx.DefaultPosition, size=wx.Size(393, 411),
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
        # startdate = self.text_startdate.GetValue().strip()
        # if len(startdate) == 0:
        #     startdate = "20170301"
        # enddate = self.text_enddate.GetValue().strip()
        # if len(enddate) == 0:
        #     enddate = int(time.strftime('%Y%m%d', time.localtime(time.time()))) + 1
        # startdate_filter = startdate[0:4] + "%2F" + startdate[4:6] + "%2F" + startdate[6:8]
        # enddate_filter = enddate[0:4] + "%2F" + enddate[4:6] + "%2F" + enddate[6:8]

        data_sn_list = []
        data_filename_original_list = []
        data_filename_final_list = []
        data_creator_list = []
        data_department_name_list = []
        data_type_invention_list = []
        data_daili_list = []
        data_management_sn_list = []
        data_shouli_sn_list = []
        data_shenqing_date_list = []



        # 模拟登陆
        url_login = "http://10.110.6.34/users/login"
        payload_login = "_method=POST&_method=POST&data%5BUser%5D%5Btype%5D=email&data%5BUser%5D%5Busername%5D={username_sub}&data%5BUser%5D%5Bpassword%5D={password_sub}".format(
            username_sub=urllib2.quote(username), password_sub=urllib2.quote(password))
        headers_base = {
            'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.8",
            'cache-control': "no-cache",
            'connection': "keep-alive",
            'content-length': "147",
            'content-type': "application/x-www-form-urlencoded",
            'host': "10.110.6.34",
            'origin': "http://10.110.6.34",
            'referer': "http://10.110.6.34/users/login",
            'upgrade-insecure-requests': "1",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
            #    'postman-token': "a99dbf7b-9cc3-8690-1f7f-fb241a97c835"
        }
        get_data = requests.session()
        get_data.post(url_login, data=payload_login, headers=headers_base)

        # 获取数据
        # 先使用limit=1来登录获取最大值。
        url_data = "http://10.110.6.34/patent/patent/index"
        #payload_1 = "filter%5BPreliminaryBase.filed_date%5D%5Bfrom%5D={starttime}&filter%5BPreliminaryBase.filed_date%5D%5Bto%5D={endtime}&limit=1&sortDirect=DESC&sortField=PreliminaryBase.filed_date".format(
        #    starttime=startdate_filter, endtime=enddate_filter)
        payload_1 = "limit=1"
        headers_data = {
            'accept': "application/json, text/javascript, */*; q=0.01",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.8",
            'connection': "keep-alive",
            'content-length': "194",
            'content-type': "application/x-www-form-urlencoded; charset=UTF-8",
            'host': "10.110.6.34",
            'origin': "http://10.110.6.34",
            'referer': "http://10.110.6.34/patent/patent/index",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
            'x-requested-with': "XMLHttpRequest",
            'cache-control': "no-cache"
        }

        response_1 = get_data.post(url_data, data=payload_1, headers=headers_data, verify=False)
        data_1 = response_1.content
        # 获取最大值
        max_number = re.search(r'"pagination":{"currentPage":1,"offset":"1","total":(\d+),', data_1).groups()[0]

        # 使用最大值来获取信息
        #payload_data = "filter%5BPreliminaryBase.filed_date%5D%5Bfrom%5D={starttime}&filter%5BPreliminaryBase.filed_date%5D%5Bto%5D={endtime}&limit={max_number}&sortDirect=DESC&sortField=PreliminaryBase.filed_date".format(
        #    max_number=max_number, starttime=startdate_filter, endtime=enddate_filter)
        payload_data = "limit={max_number}".format(max_number=max_number)
        response_data = get_data.post(url_data, data=payload_data, headers=headers_data, verify=False)
        data_original = response_data.content

        # 获取管理编号
        data_management_sn_list_temp = re.findall(r'"Patent.management_sn":"(\d+)"', data_original)

        # 获取专利类型、连接的数字、专利撰写后的名称
        # 获取总的数据
        data_link_temp = re.findall(r'"Patent.patent_name":"<span class=.*?>(.*?)<\\/span><a href=\\"http:\\/\\/10.110.6.34\\/patent\\/patent\\/view\\/(\d+)\\" target=\\"_blank\\">(.*?)<\\/a>"', data_original)
        #分别获取
        data_link_number_temp = []
        data_type_invention_list_temp = []
        data_filename_final_list_temp = []
        for item in data_link_temp:
            data_type_invention_list_temp.append(item[0].decode('unicode_escape'))
            data_link_number_temp.append(item[1])
            data_filename_final_list_temp.append(item[2].decode('unicode_escape'))
        # 再将数字连接到前置地址上
        data_link_list = ["http://10.110.6.34/patent/patent/view/" + i for i in data_link_number_temp]

        # 获取受理号
        data_shouli_sn_list_temp = re.findall(r'"PreliminaryBase.application_number":"(\w+\.*?\w*?)","PreliminaryBase.filed_date"', data_original)

        #获取申请时间
        data_shenqing_date_temp = re.findall(r'"PreliminaryBase.filed_date":"(\d+\\/\d+\\/\d+)"', data_original)
        data_shenqing_date_list_temp = [i.replace("\\/", "-") for i in data_shenqing_date_temp]


        headers_link = {
            'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.8",
            'cache-control': "no-cache",
            'connection': "keep-alive",
            'host': "10.110.6.34",
            'referer': "http://10.110.6.34/patent/patent/index",
            'upgrade-insecure-requests': "1",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36"
        }


        a = int(len(data_management_sn_list_temp) / 10)

        for index, item in enumerate(data_management_sn_list_temp):
            if index % a == 0:
                b = int(index / a) * 10
                self.updatedisplay(b)
            # sn_temp = re.search(r"\d{6}", item)
            if int(item) > 201702016313:
                data_management_sn_list.append(item)
                data_shouli_sn_list.append(data_shouli_sn_list_temp[index])
                data_shenqing_date_list.append(data_shenqing_date_list_temp[index])
                data_type_invention_list.append(data_type_invention_list_temp[index])
                data_filename_final_list.append(data_filename_final_list_temp[index])

                data_temp = get_data.get(data_link_list[index], headers=headers_link, verify=False).text
                data_soup_tobe_filter = BeautifulSoup(data_temp, "html.parser")
                name_daili = data_soup_tobe_filter.select("#patentDetail > table > tbody > tr:nth-of-type(26) > td > div > div > a")[0].get_text().strip()
                filename_original = data_soup_tobe_filter.select("#patentDetail > table > tbody > tr:nth-of-type(5) > td > div > div")[0].get_text().strip()
                name_creator_temp = data_soup_tobe_filter.select("#patentDetail > table > tbody > tr:nth-of-type(30) > td > div > div")[0].get_text().strip()
                name_creator = re.search(r"\D*", name_creator_temp).group()
                department_temp = data_soup_tobe_filter.select("#patentDetail > table > tbody > tr:nth-of-type(31) > td > div > div > a")
                department = "".join([i.get_text().strip() for i in department_temp])
                # sn_temp = data_soup_tobe_filter.select("#inventions > div > span:nth-of-type(1) > table > tbody > tr > td:nth-of-type(1) > a")[0].get_text().strip()
                data_creator_list.append(name_creator)
                data_daili_list.append(name_daili)
                data_department_name_list.append(department)
                data_filename_original_list.append(filename_original)
                # data_sn_list.append(sn_temp)


        title_sheet = ['管理编号'.decode('gbk'), '专利类型'.decode('gbk'), '原专利名称'.decode('gbk'), '代理提交专利名称'.decode('gbk'), '发明人'.decode('gbk'), '申请号'.decode('gbk'), '申请日期'.decode('gbk'), '代理'.decode('gbk'), '部门'.decode('gbk')]
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
        sheet.set_column('F:F', 18)
        sheet.set_column('G:G', 13)
        sheet.set_column('H:H', 20)
        sheet.set_column('I:I', 42)
        sheet.merge_range(0, 0, 0, 9, "%s2017财年受理专利总览".decode('gbk') % department_write, formattitle)
        for index_title, item_title in enumerate(title_sheet):
            sheet.write(1, index_title, item_title, formatone)
        for index_data, item_data in enumerate(data_management_sn_list):
            sheet.write(2 + index_data, 0, item_data, formatone)
            sheet.write(2 + index_data, 1, data_type_invention_list[index_data], formatone)
            sheet.write(2 + index_data, 2, data_filename_original_list[index_data], formatone)
            #sheet.write(2 + index_data, 3, data_sn_list[index_data], formatone)
            sheet.write(2 + index_data, 3, data_filename_final_list[index_data], formatone)
            sheet.write(2 + index_data, 4, data_creator_list[index_data], formatone)
            sheet.write(2 + index_data, 5, data_shouli_sn_list_temp[index_data], formatone)
            sheet.write_datetime(2 + index_data, 6, datetime.datetime.strptime(data_shenqing_date_list_temp[index_data],
                                                                               '%Y-%m-%d'),
                                 workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))

            sheet.write(2 + index_data, 7, data_daili_list[index_data], formatone)
            sheet.write(2 + index_data, 8, data_department_name_list[index_data], formatone)
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
            self.output_info.AppendText("完成".decode('gbk') + unicode(t) + "%")
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
