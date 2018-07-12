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
from multiprocessing import Pool
import multiprocessing


def getpage(page_number, startdate_filter, enddate_filter, get_data):
    list_status_temp = []
    list_current_node_temp = []
    list_num_temp = []
    list_sn_filename_temp = []
    list_assign_temp = []
    list_date_created_temp = []
    list_creator_temp = []
    headers_data = {
        'host': "10.110.6.34",
        'connection': "keep-alive",
        'content-length': "53",
        'accept': "application/json, text/javascript, */*; q=0.01",
        'origin': "http://10.110.6.34",
        'x-requested-with': "XMLHttpRequest",
        'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
        'content-type': "application/x-www-form-urlencoded; charset=UTF-8",
        'referer': "http://10.110.6.34/audit/bpm/complete",
        'accept-encoding': "gzip, deflate",
        'accept-language': "zh-CN,zh;q=0.8",
    }
    url_data = "http://10.110.6.34/audit/bpm/complete"

    payload_data = "page={page}&filter%5Bcreated_at%5D%5Bfrom%5D={start_date}&filter%5Bcreated_at%5D%5Bto%5D={end_date}&limit=20&sortDirect=DESC&sortField=created_at".format(
        page=page_number, start_date=startdate_filter, end_date=enddate_filter)
    response_data = get_data.post(url_data, data=payload_data, headers=headers_data, verify=False)
    return_code_page = response_data.status_code
    data_original = ''
    print(str(page_number) + " " + str(return_code_page))
    if return_code_page != 200:
        print("Try to reget page info for page %s" % str(page_number))
        for i in range(1, 10):
            response_data_try = get_data.post(url_data, data=payload_data, headers=headers_data, verify=False)
            print("Try %s times for page %s" % (str(i), str(page_number)))
            if response_data_try.status_code != 200:
                continue
            else:
                data_original = response_data_try.content
                print(str(page_number) + " " + str(response_data_try.status_code))
                break
    else:
        data_original = response_data.content
    # print data_original
    # 获取状态
    list_status_temp_1 = re.findall(r'"status":"<span class=my_task_status_\w*?>(.*?)<\\/span>', data_original)
    list_status_temp_2 = [item.decode('unicode_escape') for item in list_status_temp_1]
    # print list_status_temp_2
    # 获取提交时间
    list_date_created_temp_1 = re.findall(r'"created_at":"(\d+\\/\d+\\/\d+)', data_original)
    list_date_created_temp_2 = [item.replace("\\/", "-") for item in list_date_created_temp_1]

    # 获取撰写人
    list_creator_temp_1 = re.findall(r',"created_by":"(.*?)",', data_original)
    list_creator_temp_3 = [item.decode('unicode_escape') for item in list_creator_temp_1 if re.search(r'<', item) is None]
    # print list_creator_temp_4
    list_creator_temp_2 = list_creator_temp_3

    # 获取当前处理人
    list_assign_temp_1 = re.findall(r'"assignee":"(.*?)"', data_original)
    list_assign_temp_1_temp = list_assign_temp_1[1:]
    list_assign_temp_2 = [item.decode('unicode_escape') for item in list_assign_temp_1_temp]
    # print list_creator_temp_1
    # 获取当前处理节点
    list_current_node_temp_1 = re.findall(
        r'"node_name":"<span class=\\"node-icon\\" data-bind=\w*? title=.*?><\\/span>(.*?)",', data_original)
    list_current_node_temp_2 = [item.decode('unicode_escape') for item in list_current_node_temp_1]
    # print  list_current_node_temp_2
    # 获取链接和编号和文件名
    list_data_sn_filename_temp = re.findall(
        r'"subject":"<a href=\\"\\/invention\\/inventions\\/view\\/(\d+)\\" target=\\"_blank\\">(.*?)<\\/a>",',
        data_original)
    list_num_temp_1 = []
    list_sn_filename_temp_1 = []
    for item in list_data_sn_filename_temp:
        list_num_temp_1.append(item[0])
        list_sn_filename_temp_1.append(item[1].decode('unicode_escape'))

    list_status_temp.extend(list_status_temp_2)
    # list_creator_temp.extend(list_creator_temp_2)
    list_date_created_temp.extend(list_date_created_temp_2)
    list_current_node_temp.extend(list_current_node_temp_2)
    list_num_temp.extend(list_num_temp_1)
    list_sn_filename_temp.extend(list_sn_filename_temp_1)
    list_assign_temp.extend(list_assign_temp_2)
    list_creator_temp.extend(list_creator_temp_2)
    return list_status_temp, list_date_created_temp, list_current_node_temp, list_num_temp, list_sn_filename_temp, list_assign_temp, list_creator_temp


def getdetail(link, applicant_link, get_data):
    list_status_second_except = ["撰写驳回".decode('gbk'), "待决定".decode('gbk')]
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
    print link
    data_detail_temp = get_data.get(link, headers=headers_link, verify=False)
    if data_detail_temp.status_code != 404:
        data_temp = data_detail_temp.text
        data_soup_tobe_filter = BeautifulSoup(data_temp, "html.parser")
        # print data_soup_tobe_filter
        # print data_soup_tobe_filter
        status_second = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(3) > td")[0].get_text().strip()
        if status_second not in list_status_second_except:
            # 专利类型
            type_invention = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(6) > td")[0].get_text().strip()
            # 撰写人
            # creator_temp = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(9) > td")[0].get_text().strip()
            # creator = re.search(r"\D*", creator_temp).group()
            # 部门
            department_temp = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(10) > td > a")
            department = "".join([i.get_text().strip() for i in department_temp])
            # 代理机构
            name_daili_department = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(14) > td > a")
            # 代理人
            name_daili_person = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(15) > td > a")

            # 最后更新人
            username_last_update_temp = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(22) > td")[0].get_text().strip().split(" ")[0]
            username_last_update = re.search(r"\D*", username_last_update_temp).group()

            # 最后更新时间
            date_last_update_temp = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(22) > td")[0].get_text().strip().split(" ")[1]
            date_last_update = date_last_update_temp.replace("/", "-")

            if len(name_daili_department) != 0:
                data_daili_department = name_daili_department[0].get_text().strip()
            else:
                data_daili_department = "None"

            if len(name_daili_person) != 0:
                data_daili_person = name_daili_person[0].get_text().strip()
            else:
                data_daili_person = "None"
            # 获取申请人信息
            data_applicant_temp = get_data.get(applicant_link, headers=headers_link, verify=False).text
            print applicant_link
            applicant_info_temp = re.search(r'"assignee":"(.*?)",', data_applicant_temp)
            if applicant_info_temp is not None:
                applicant_info = (applicant_info_temp.groups()[0]).decode('unicode_escape')
            else:
                applicant_info = "None"
            # applicant_info = applicant_info_temp.decode('unicode_escape')
            return link, type_invention, username_last_update, date_last_update, status_second, department, applicant_info, data_daili_department, data_daili_person
    else:
        return "None"


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
            self.output_info.AppendText("完成".decode('gbk') + unicode(t) + "%")
        elif t == "Finished":
            self.button_go.Enable()
        else:
            self.output_info.AppendText("%s".decode('gbk') % t)
        self.output_info.AppendText(os.linesep)

    def run(self):
        self.updatedisplay("开始抓取".decode('gbk'))
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        department_write = self.input_department.GetValue()
        username = self.input_username.GetValue()
        password = self.input_password.GetValue()
        startdate = self.text_startdate.GetValue().strip()
        list_status_second_except = ["撰写驳回".decode('gbk'), "待决定".decode('gbk')]
        if len(startdate) == 0:
            startdate = "20180320"
        enddate = self.text_enddate.GetValue().strip()
        if len(enddate) == 0:
            enddate = int(time.strftime('%Y%m%d', time.localtime(time.time()))) + 1
        startdate_filter = startdate[0:4] + "%2F" + startdate[4:6] + "%2F" + startdate[6:8]
        # print startdate_filter
        enddate_filter = enddate[0:4] + "%2F" + enddate[4:6] + "%2F" + enddate[6:8]
        # print enddate_filter
        # 排除在外的状态需要特殊考虑，有可能出现显示这些特殊状态，但是实际是受理的！
        list_except = ["驳回".decode('gbk'), '退回发起人'.decode('gbk'), '撤销'.decode('gbk')]
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
        }
        get_data = requests.session()
        get_data.post(url_login, data=payload_login, headers=headers_base)
        # 获取数据
        # 先使用limit=1来登录获取最大值。
        url_data = "http://10.110.6.34/audit/bpm/complete"
        payload_1 = "filter%5Bcreated_at%5D%5Bfrom%5D={start_date}&filter%5Bcreated_at%5D%5Bto%5D={end_date}&limit=1&sortDirect=DESC&sortField=created_at".format(start_date=startdate_filter, end_date=enddate_filter)
        headers_data = {
            'host': "10.110.6.34",
            'connection': "keep-alive",
            'content-length': "53",
            'accept': "application/json, text/javascript, */*; q=0.01",
            'origin': "http://10.110.6.34",
            'x-requested-with': "XMLHttpRequest",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
            'content-type': "application/x-www-form-urlencoded; charset=UTF-8",
            'referer': "http://10.110.6.34/audit/bpm/complete",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.8",
            }

        response_1 = get_data.post(url_data, data=payload_1, headers=headers_data, verify=False)
        data_1 = response_1.content
        # 获取最大值
        max_number = re.search(r'"pagination":{"currentPage":1,"offset":"1","total":(\d+),', data_1).groups()[0]
        print max_number
        # 无法直接使用最大值获取，分页获取然后合并
        total_page = int(max_number) / 20 + 2
        # print total_page
        # 分页来获取信息
        self.updatedisplay("开始抓取第一部分，分页总数据！".decode('gbk'))
        # use multiprocessing to get page data
        temp = []
        pool_page = Pool()
        for page_number in range(1, total_page):
            temp.append(pool_page.apply_async(getpage, args=(page_number, startdate_filter, enddate_filter, get_data)))
            self.updatedisplay("已抓取%s/%s页！".decode('gbk') % (page_number, total_page - 1))
        pool_page.close()
        pool_page.join()
        print("1")
        # filter page data
        list_status_temp_after_pagedata_get = []
        list_date_created_temp_after_pagedata_get = []
        list_current_node_temp_after_pagedata_get = []
        list_num_temp_after_pagedata_get = []
        list_sn_filename_temp_after_pagedata_get = []
        list_assign_temp_after_pagedata_get = []
        list_creator_temp_after_pagedata_get = []

        for item_page in temp:
            data_temp = item_page.get()
            list_status_temp_after_pagedata_get.extend(data_temp[0])
            list_date_created_temp_after_pagedata_get.extend(data_temp[1])
            list_current_node_temp_after_pagedata_get.extend(data_temp[2])
            list_num_temp_after_pagedata_get.extend(data_temp[3])
            list_sn_filename_temp_after_pagedata_get.extend(data_temp[4])
            list_assign_temp_after_pagedata_get.extend(data_temp[5])
            list_creator_temp_after_pagedata_get.extend(data_temp[6])
        print len(list_sn_filename_temp_after_pagedata_get)
        print len(list_creator_temp_after_pagedata_get)

        # 先处理一遍数据，把名字被删除只剩下\/的、SN重复的去除、撰写人为徐莉（浪潮信息）和王文晓的去掉
        list_status = []
        list_date_created = []
        list_current_node = []
        list_sn_filename = []
        list_num = []
        list_sn = []
        list_filename = []
        list_assign = []
        list_creator = []

        # 特殊状态的列表
        list_status_special = []
        list_date_created_special = []
        list_current_node_special = []
        list_sn_filename_special = []
        list_num_special = []
        list_sn_special = []
        list_filename_special = []
        list_assign_special = []
        list_creator_special = []

        # 排除撰写人为专利工程师的
        list_creator_except = ["徐莉（浪潮信息）".decode('gbk'), "王文晓".decode('gbk'), "张钟元".decode('gbk')]
        # 用于排除被驳回之后到开始节点的
        list_except_current_node = ["开始节点".decode('gbk')]

        # 先获取非特殊列表里面的，然后再获取在特殊列表里面的，如果在特殊列表中的SN但是不在非特殊列表的结果中，那就需要增加进去。
        # 获取非特殊列表中的
        for index_status, item_status in enumerate(list_status_temp_after_pagedata_get):
            if item_status not in list_except and list_sn_filename_temp_after_pagedata_get[index_status] != '\/' and list_creator_temp_after_pagedata_get[index_status] not in list_creator_except:
                if list_sn_filename_temp_after_pagedata_get[index_status].split("\\/")[0] not in list_sn:
                    list_sn.append(list_sn_filename_temp_after_pagedata_get[index_status].split("\\/")[0])
                    list_filename.append(list_sn_filename_temp_after_pagedata_get[index_status].split("\\/")[1])
                    list_status.append(item_status)
                    # list_creator.append(list_creator_temp[index_status])
                    list_date_created.append(list_date_created_temp_after_pagedata_get[index_status])
                    list_current_node.append(list_current_node_temp_after_pagedata_get[index_status])
                    list_sn_filename.append(list_sn_filename_temp_after_pagedata_get[index_status])
                    list_num.append(list_num_temp_after_pagedata_get[index_status])
                    list_assign.append(list_assign_temp_after_pagedata_get[index_status])
                    list_creator.append(list_creator_temp_after_pagedata_get[index_status])
        #获取特殊列表中的
        for index_status_special, item_status_special in enumerate(list_status_temp_after_pagedata_get):
            if item_status_special in list_except and list_sn_filename_temp_after_pagedata_get[index_status_special] != '\/' and list_current_node_temp_after_pagedata_get[index_status_special] not in list_except_current_node and list_creator_temp_after_pagedata_get[index_status_special] not in list_creator_except:
                if list_sn_filename_temp_after_pagedata_get[index_status_special].split("\\/")[0] not in list_sn:
                    list_sn_special.append(list_sn_filename_temp_after_pagedata_get[index_status_special].split("\\/")[0])
                    list_filename_special.append(list_sn_filename_temp_after_pagedata_get[index_status_special].split("\\/")[1])
                    list_status_special.append(item_status_special)
                    # list_creator.append(list_creator_temp[index_status])
                    list_date_created_special.append(list_date_created_temp_after_pagedata_get[index_status_special])
                    list_current_node_special.append(list_current_node_temp_after_pagedata_get[index_status_special])
                    list_sn_filename_special.append(list_sn_filename_temp_after_pagedata_get[index_status_special])
                    list_num_special.append(list_num_temp_after_pagedata_get[index_status_special])
                    list_assign_special.append(list_assign_temp_after_pagedata_get[index_status_special])
                    list_creator_special.append(list_creator_temp_after_pagedata_get[index_status_special])
        # 开始获取每个专利的信息
        #获取链接，分为普通的和特殊的
        list_link = ["http://10.110.6.34/invention/inventions/view/" + i for i in list_num]
        list_link_special = ["http://10.110.6.34/invention/inventions/view/" + i for i in list_num_special]
        # 申请人信息的链接，分为普通的和特殊的
        list_applicant_link = ["http://10.110.6.34/invention/async_invention_applicant/async_list/" + i for i in list_num]
        list_applicant_link_special = ["http://10.110.6.34/invention/async_invention_applicant/async_list/" + i for i in list_num_special]

        # 已字典的形式保存信息，达到跟前面能够对应的接口,通过list_num来做接口
        # 获取普通专利的状态
        self.updatedisplay("开始抓取第二部分！每个专利的信息！".decode('gbk'))
        self.updatedisplay("开始抓取正常流程的专利的信息！".decode('gbk'))
        dict_data_detail = {}
        for item_1 in list_num:
            dict_data_detail["%s" % item_1] = []
        temp_detail = []
        pool_detail = Pool()
        for index, item_2 in enumerate(list_link):
            applicant_link = list_applicant_link[index]
            # return link, type_invention, username_last_update, date_last_update, status_second, department, creator, applicant_info, data_daili_department, data_daili_person
            temp_detail.append(pool_detail.apply_async(getdetail, args=(item_2, applicant_link, get_data)))
        pool_detail.close()
        pool_detail.join()
        # 处理获取的普通专利的每个专利的信息
        # return link, type_invention, username_last_update, date_last_update, status_second,
        # department, applicant_info, data_daili_department, data_daili_person
        for item_detail in temp_detail:
            data_detail_temp = item_detail.get()
            if data_detail_temp is not None:
                if data_detail_temp[0] != "None":
                    num = data_detail_temp[0].split("/")[-1]
                    index_to_log = list_num.index(num)
                    # status_second
                    dict_data_detail["%s" % num].append(data_detail_temp[4])
                    # sn
                    dict_data_detail["%s" % num].append(list_sn[index_to_log])
                    # filename
                    dict_data_detail["%s" % num].append(list_filename[index_to_log])
                    # department
                    dict_data_detail["%s" % num].append(data_detail_temp[5])
                    # type_invention
                    dict_data_detail["%s" % num].append(data_detail_temp[1])
                    # creator
                    dict_data_detail["%s" % num].append(list_creator[index_to_log])
                    # create_date
                    dict_data_detail["%s" % num].append(list_date_created[index_to_log])
                    # username_lastupdate
                    dict_data_detail["%s" % num].append(data_detail_temp[2])
                    # lastupdate_date
                    dict_data_detail["%s" % num].append(data_detail_temp[3])
                    # current_node
                    dict_data_detail["%s" % num].append(list_current_node[index_to_log])
                    # name_daili_department
                    dict_data_detail["%s" % num].append(data_detail_temp[7])
                    # name_daili
                    dict_data_detail["%s" % num].append(data_detail_temp[8])
                    # assign
                    dict_data_detail["%s" % num].append(list_assign[index_to_log])
                    # applicant
                    dict_data_detail["%s" % num].append(data_detail_temp[6])

        # 获取特殊专利的状态
        self.updatedisplay("开始抓取异常流程专利的信息！".decode('gbk'))
        dict_data_detail_special = {}
        for item_3 in list_num_special:
            dict_data_detail_special["%s" % item_3] = []
        temp_detail_special = []
        pool_detail_special = Pool()
        for index_special, item_special in enumerate(list_link_special):
            applicant_link_special = list_applicant_link_special[index_special]
            # return link, type_invention, username_last_update, date_last_update, status_second, department, creator, applicant_info, data_daili_department, data_daili_person
            temp_detail_special.append(pool_detail_special.apply_async(getdetail, args=(item_special, applicant_link_special, get_data)))
        pool_detail_special.close()
        pool_detail_special.join()
        # 处理获取的特殊专利的每个专利的信息
        for item_detail_special in temp_detail_special:
            data_detail_temp_special = item_detail_special.get()
            if data_detail_temp_special is not None:
                if data_detail_temp_special[0] != "None":
                    num_special = data_detail_temp_special[0].split("/")[-1]
                    index_to_log_special = list_num_special.index(num_special)
                    # status_second
                    dict_data_detail_special["%s" % num_special].append(data_detail_temp_special[4])
                    # sn
                    dict_data_detail_special["%s" % num_special].append(list_sn_special[index_to_log_special])
                    # filename
                    dict_data_detail_special["%s" % num_special].append(list_filename_special[index_to_log_special])
                    # department
                    dict_data_detail_special["%s" % num_special].append(data_detail_temp_special[5])
                    # type_invention
                    dict_data_detail_special["%s" % num_special].append(data_detail_temp_special[1])
                    # creator
                    dict_data_detail_special["%s" % num_special].append(list_creator_special[index_to_log_special])
                    # create_date
                    dict_data_detail_special["%s" % num_special].append(list_date_created_special[index_to_log_special])
                    # username_lastupdate
                    dict_data_detail_special["%s" % num_special].append(data_detail_temp_special[2])
                    # lastupdate_date
                    dict_data_detail_special["%s" % num_special].append(data_detail_temp_special[3])
                    # current_node
                    dict_data_detail_special["%s" % num_special].append(list_current_node_special[index_to_log_special])
                    # name_daili_department
                    dict_data_detail_special["%s" % num_special].append(data_detail_temp_special[7])
                    # name_daili
                    dict_data_detail_special["%s" % num_special].append(data_detail_temp_special[8])
                    # assign
                    dict_data_detail_special["%s" % num_special].append(list_assign_special[index_to_log_special])
                    # applicant
                    dict_data_detail_special["%s" % num_special].append(data_detail_temp_special[6])
        # 处理数据,将普通的和特殊的合并到一起
        list_status_second_write = []
        list_sn_write = []
        list_filename_write = []
        list_department_write = []
        list_type_write = []
        list_creator_write = []
        list_date_created_write = []
        list_username_lastupdate_write = []
        list_date_lastupdate_write = []
        list_current_node_write = []
        list_name_daili_department_write = []
        list_name_daili_person_write = []
        list_assign_write = []
        list_applicant_write = []
        #普通的
        # list_status_second_except = ["撰写驳回".decode('gbk'), "待决定".decode('gbk')]

        for item_putong in dict_data_detail:
            if len(dict_data_detail[item_putong]) != 0:
                if dict_data_detail[item_putong][0] not in list_status_second_except:
                    if dict_data_detail[item_putong][0] == "撰写中".decode('gbk') and dict_data_detail[item_putong][10] == "None":
                        continue
                    list_status_second_write.append(dict_data_detail[item_putong][0])
                    list_sn_write.append(dict_data_detail[item_putong][1])
                    list_filename_write.append(dict_data_detail[item_putong][2])
                    list_department_write.append(dict_data_detail[item_putong][3])
                    list_type_write.append(dict_data_detail[item_putong][4])
                    list_creator_write.append(dict_data_detail[item_putong][5])
                    list_date_created_write.append(dict_data_detail[item_putong][6])
                    list_username_lastupdate_write.append(dict_data_detail[item_putong][7])
                    list_date_lastupdate_write.append(dict_data_detail[item_putong][8])
                    list_current_node_write.append(dict_data_detail[item_putong][9])
                    list_name_daili_department_write.append(dict_data_detail[item_putong][10])
                    list_name_daili_person_write.append(dict_data_detail[item_putong][11])
                    list_assign_write.append(dict_data_detail[item_putong][12])
                    list_applicant_write.append(dict_data_detail[item_putong][13])
        # 特殊的
        for item_special in dict_data_detail_special:
            if len(dict_data_detail_special[item_special]) != 0:
                if dict_data_detail_special[item_special][0] not in list_status_second_except:
                    # 排除掉撰写中状态但是代理信息却为空的专利。此种专利为发明人自行发起撰写流程，需要排除！
                    if dict_data_detail_special[item_special][0] == "撰写中".decode('gbk') and dict_data_detail_special[item_special][10] == "None":
                        continue
                    if dict_data_detail_special[item_special][0] == "提案中".decode('gbk'):
                        continue
                    list_status_second_write.append(dict_data_detail_special[item_special][0])
                    list_sn_write.append(dict_data_detail_special[item_special][1])
                    list_filename_write.append(dict_data_detail_special[item_special][2])
                    list_department_write.append(dict_data_detail_special[item_special][3])
                    list_type_write.append(dict_data_detail_special[item_special][4])
                    list_creator_write.append(dict_data_detail_special[item_special][5])
                    list_date_created_write.append(dict_data_detail_special[item_special][6])
                    list_username_lastupdate_write.append(dict_data_detail_special[item_special][7])
                    list_date_lastupdate_write.append(dict_data_detail_special[item_special][8])
                    list_current_node_write.append(dict_data_detail_special[item_special][9])
                    list_name_daili_department_write.append(dict_data_detail_special[item_special][10])
                    list_name_daili_person_write.append(dict_data_detail_special[item_special][11])
                    list_assign_write.append(dict_data_detail_special[item_special][12])
                    list_applicant_write.append(dict_data_detail_special[item_special][13])

        print "sn length " + str(len(list_sn_write))
        print "status length " + str(len(list_status_second_write))
        print "current node length " + str(len(list_current_node_write))
        print "creator length " + str(len(list_creator_write))
        print "date created length" + str(len(list_date_created_write))
        print "filename length " + str(len(list_filename_write))
        print "username lastupdate length " + str(len(list_username_lastupdate_write))
        print "date lastupdate length " + str(len(list_date_lastupdate_write))
        print "department length " + str(len(list_department_write))
        print "type length " + str(len(list_type_write))
        print "daili department length " + str(len(list_name_daili_department_write))
        print "daili person length " + str(len(list_name_daili_person_write))
        print "assign length " + str(len(list_assign_write))
        print "applicant list length " + str(len(list_applicant_write))

        # write info to excel file
        title_sheet = ['当前状态'.decode('gbk'), '提案编号'.decode('gbk'), '提案名称'.decode('gbk'), '处别'.decode('gbk'),
                       '专利类型'.decode('gbk'), '撰写人'.decode('gbk'), '提交时间'.decode('gbk'), '最后更新人'.decode('gbk'),
                       '最后更新时间'.decode('gbk'), '当前节点'.decode('gbk'), '代理机构名称'.decode('gbk'), '代理人'.decode('gbk'),
                       '当前处理人'.decode('gbk'), '申请人'.decode('gbk')]
        timestamp = time.strftime('%Y%m%d', time.localtime())
        # department_write = "测试验证部".decode('gbk')
        workbook_display = xlsxwriter.Workbook('2018财年%s专利总览-%s.xlsx'.decode('gbk') % (department_write, timestamp))
        sheet = workbook_display.add_worksheet('2018财年%s专利总览'.decode('gbk') % department_write)
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
        sheet.set_column('F:F', 11)
        sheet.set_column('G:I', 15)
        sheet.set_column('J:J', 17)
        sheet.set_column('K:L', 33)
        sheet.set_column('M:M', 20)
        sheet.set_column('N:N', 25)

        sheet.merge_range(0, 0, 0, 13, "%s2018财年专利总览".decode('gbk') % department_write, formattitle)
        for index_title, item_title in enumerate(title_sheet):
            sheet.write(1, index_title, item_title, formatone)
        for index_data, item_data in enumerate(list_status_second_write):
            if item_data not in list_status_second_except:
                sheet.write(2 + index_data, 0, item_data, formatone)
                sheet.write(2 + index_data, 1, list_sn_write[index_data], formatone)
                sheet.write(2 + index_data, 2, list_filename_write[index_data], formatone)
                sheet.write(2 + index_data, 3, list_department_write[index_data], formatone)
                sheet.write(2 + index_data, 4, list_type_write[index_data], formatone)
                sheet.write(2 + index_data, 5, list_creator_write[index_data], formatone)
                sheet.write_datetime(2 + index_data, 6,
                                     datetime.datetime.strptime(list_date_created_write[index_data], '%Y-%m-%d'),
                                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))

                sheet.write(2 + index_data, 7, list_username_lastupdate_write[index_data], formatone)
                sheet.write_datetime(2 + index_data, 8,
                                     datetime.datetime.strptime(list_date_lastupdate_write[index_data], '%Y-%m-%d'),
                                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
                sheet.write(2 + index_data, 9, list_current_node_write[index_data], formatone)
                sheet.write(2 + index_data, 10, list_name_daili_department_write[index_data], formatone)
                sheet.write(2 + index_data, 11, list_name_daili_person_write[index_data], formatone)
                if len(list_assign_write[index_data]) == 0:
                    sheet.write(2 + index_data, 12, None, formatone)
                else:
                    sheet.write(2 + index_data, 12, list_assign_write[index_data], formatone)
                sheet.write(2 + index_data, 13, list_applicant_write[index_data], formatone)

        workbook_display.close()
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        self.updatedisplay("抓取结束,请点击退出按钮退出程序".decode('gbk'))
        time.sleep(1)
        self.updatedisplay("Finished")


if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = wx.App()
    frame = FrameZhuanli(None)
    frame.Show()
    app.MainLoop()
