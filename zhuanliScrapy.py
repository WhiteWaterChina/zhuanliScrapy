#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com
import requests
import re
from bs4 import BeautifulSoup
import xlsxwriter
import os
import time
import datetime
from threading import Thread
import wx
import urllib2

class FrameZhuanli(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"专利信息获取系统", pos=wx.DefaultPosition, size=wx.Size(393, 411),
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
        startdate_filter = startdate[0:4] + "%2F" + startdate[4:6] + "%2F" + startdate[6:8]
        enddate_filter = enddate[0:4] + "%2F" + enddate[4:6] + "%2F" + enddate[6:8]
        #排除在外的状态
        except_list = ["撰写驳回".decode('gbk'),'待决定'.decode('gbk')]
        #模拟登陆
        url_login = "http://10.110.6.34/users/login"
        payload_login = "_method=POST&_method=POST&data%5BUser%5D%5Btype%5D=email&data%5BUser%5D%5Busername%5D={username_sub}&data%5BUser%5D%5Bpassword%5D={password_sub}".format(username_sub=urllib2.quote(username), password_sub=urllib2.quote(password))
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

        #获取数据
        #先使用limit=1来登录获取最大值。
        url_data = "http://10.110.6.34/invention/inventions/index"
        payload_1 = "limit=1&filter%5BInvention.updated%5D%5Bfrom%5D={starttime}&filter%5BInvention.updated%5D%5Bto%5D={endtime}".format(starttime=startdate_filter, endtime=enddate_filter)
        headers_data = {
            'host': "10.110.6.34",
            'connection': "keep-alive",
            'content-length': "29",
            'accept': "application/json, text/javascript, */*; q=0.01",
            'origin': "http://10.110.6.34",
            'x-requested-with': "XMLHttpRequest",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
            'content-type': "application/x-www-form-urlencoded",
            'referer': "http://10.110.6.34/invention/inventions/index",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.8",
            'cache-control': "no-cache",
            }

        response_1 = get_data.post(url_data, data=payload_1, headers=headers_data, verify=False)
        data_1 = response_1.content
        #获取最大值
        max_number = re.search(r'"pagination":{"currentPage":1,"offset":"1","total":(\d+),', data_1).groups()[0]

        #使用最大值来获取信息
        #payload_data = "limit=%s" % max_number
        payload_data = "limit={max_number}&filter%5BInvention.updated%5D%5Bfrom%5D={starttime}&filter%5BInvention.updated%5D%5Bto%5D={endtime}".format(max_number=max_number,starttime=startdate_filter, endtime=enddate_filter)
        response_data = get_data.post(url_data, data=payload_data, headers=headers_data, verify=False)
        data_original = response_data.content

        #获取编号
        list_data_sn = re.findall(r'\"Invention.track_number\"\:\"(\d+)"', data_original)
        #获取链接
        #获取链接的数字
        data_link_temp = re.findall(r'"Invention.title":"<a href=\\"http:\\/\\/10.110.6.34\\/invention\\/inventions\\/view\\/(\d+)\\" target=\\"_blank\\"', data_original)
        #再将数字连接到前置地址上
        list_data_link = ["http://10.110.6.34/invention/inventions/view/" + i for i in data_link_temp]
        #获取专利名称。先获取返回值，然后再转换编码
        data_name_temp = re.findall(r'"Invention.title":"<a.*?target=\\"_blank\\">(.*?)<\\/a>', data_original)
        list_data_name = [i.decode('unicode_escape') for i in data_name_temp]
        #获取部门和处。先获取返回值，然后再处理编码和替换多余字符
        data_department_temp = re.findall(r'"Invention.organization":"<a.*?title=(.*?)>', data_original)
        list_data_department = [i.decode('unicode_escape').replace(" &gt; ", "") for i in data_department_temp]
        #获取创建时间。先获取返回值，然后替换字符
        data_created_date_temp = re.findall(r'"Invention.created":"(\d+\\/\d+\\/\d+)"', data_original)
        list_data_created_date = [i.replace("\\/", "-") for i in data_created_date_temp]
        #获取更新时间。先获取返回值，然后替换字符
        data_update_date_temp = re.findall(r'"Invention.updated":"(\d+\\/\d+\\/\d+)"', data_original)
        list_data_update_date = [i.replace("\\/", "-") for i in data_update_date_temp]
        #获取当前状态.先获取返回值，然后再转换编码
        data_status_temp = re.findall(r'"Invention.node_status":"<a href=.*?>(.*?)<\\/a>', data_original)
        list_data_status = [i.decode('unicode_escape') for i in data_status_temp]
        #先处理一遍数据，把撰写驳回或者加上待决定的去除
        list_status = []
        list_sn = []
        list_link = []
        list_name = []
        list_department = []
        list_created_date = []
        list_updated_date = []
        for index_status, item_status in enumerate(list_data_status):
            if item_status not in except_list:
                list_status.append(item_status)
                list_sn.append(list_data_sn[index_status])
                list_link.append(list_data_link[index_status])
                list_name.append(list_data_name[index_status])
                list_department.append(list_data_department[index_status])
                list_created_date.append(list_data_created_date[index_status])
                list_updated_date.append(list_data_update_date[index_status])
        headers_link = {
            'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.8",
            'cache-control': "no-cache",
            'connection': "keep-alive",
            'host': "10.110.6.34",
            'referer': "http://10.110.6.34/invention/inventions/index",
            'upgrade-insecure-requests': "1",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
            }

        list_data_daili = []
        list_data_name_lastupdate = []
        list_type_invention = []
        list_username_created = []

        a = int(len(list_status) / 10)

        for index, item in enumerate(list_link):
            if index % a == 0:
                b = int(index / a) * 10
                self.updatedisplay(b)
            data_temp = get_data.get(item, headers=headers_link, verify=False).text
            data_soup_tobe_filter = BeautifulSoup(data_temp, "html.parser")
            type_invention = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(6) > td")[0].get_text().strip()
            name_daili = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(14) > td > a")
            name_last_update_temp = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(21) > td")[0].get_text().strip().split(" ")[0]
            name_last_update = re.search(r"\D*", name_last_update_temp).group()
            name_creator_temp = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(20) > td")[0].get_text().strip().split(" ")[0]
            name_creator = re.search(r"\D*", name_creator_temp).group()


            if len(name_daili) != 0:
                list_data_daili.append(name_daili[0].get_text().strip())
            else:
                list_data_daili.append("None")
            list_type_invention.append(type_invention)
            list_data_name_lastupdate.append(name_last_update)
            list_username_created.append(name_creator)

        title_sheet = ['当前状态'.decode('gbk'), '提案编号'.decode('gbk'), '提案名称'.decode('gbk'), '处别'.decode('gbk'),
                       '专利类型'.decode('gbk'), '撰写人'.decode('gbk'), '创建时间'.decode('gbk'),
                       '最后更新人'.decode('gbk'), '最后更新时间'.decode('gbk'), '代理名称'.decode('gbk')]
        timestamp = time.strftime('%Y%m%d', time.localtime())
        workbook_display = xlsxwriter.Workbook('%s专利总览-%s.xlsx'.decode('gbk') % (department_write, timestamp))
        sheet = workbook_display.add_worksheet('2017财年%s专利总览'.decode('gbk') % department_write)
        formatone = workbook_display.add_format()
        formatone.set_border(1)
        formattwo = workbook_display.add_format()
        formattwo.set_border(1)
        formattitle = workbook_display.add_format()
        formattitle.set_border(1)
        formattitle.set_align('center')
        formattitle.set_bg_color("yellow")
        formattitle.set_bold(True)

        sheet.set_column('B:B', 14)
        sheet.set_column('C:C', 42)
        sheet.set_column('D:D', 33)
        sheet.set_column('G:I', 15)
        sheet.set_column('J:J', 17)

        sheet.set_column('K:L', 14)
        sheet.merge_range(0, 0, 0, 9, "%s2017财年专利总览".decode('gbk') % department_write, formattitle)
        for index_title, item_title in enumerate(title_sheet):
            sheet.write(1, index_title, item_title, formatone)
            for index_data, item_data in enumerate(list_status):
                sheet.write(2 + index_data, 0, item_data, formatone)
                sheet.write(2 + index_data, 1, list_sn[index_data], formatone)
                sheet.write(2 + index_data, 2, list_name[index_data], formatone)
                sheet.write(2 + index_data, 3, list_department[index_data], formatone)
                sheet.write(2 + index_data, 4, list_type_invention[index_data], formatone)
                sheet.write(2 + index_data, 5, list_username_created[index_data], formatone)
                sheet.write_datetime(2 + index_data, 6, datetime.datetime.strptime(list_created_date[index_data],
                                                                                   '%Y-%m-%d'),
                                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))

                sheet.write(2 + index_data, 7, list_data_name_lastupdate[index_data], formatone)
                sheet.write_datetime(2 + index_data, 8, datetime.datetime.strptime(list_data_update_date[index_data],
                                                                                    '%Y-%m-%d'),
                                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
                sheet.write(2 + index_data, 9, list_data_daili[index_data], formatone)
        workbook_display.close()
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        self.updatedisplay("抓取结束,请点击退出按钮退出程序".decode('gbk'))
        time.sleep(1)
        self.updatedisplay("Finished")

    def close(self, event):
        self.Close()

    def onbutton(self, event):
        self._thread.start()
        self.started = True
        self.button_go = event.GetEventObject()
        self.button_go.Disable()

    def updatedisplay(self, msg):
        t = msg
        if isinstance(t, int):
            self.output_info.AppendText("完成".decode('gbk') + unicode(t) + "%" )
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