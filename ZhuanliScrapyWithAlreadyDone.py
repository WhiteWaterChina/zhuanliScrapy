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
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"ר����Ϣ��ȡϵͳ", pos=wx.DefaultPosition, size=wx.Size(393, 411),
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

        self.text_department = wx.StaticText(self.m_panel1, wx.ID_ANY, u"�����벿������", wx.DefaultPosition, wx.Size(150, 20),
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

        self.text_username = wx.StaticText(self.m_panel1, wx.ID_ANY, u"�������û���", wx.DefaultPosition, wx.Size(150, 20),
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

        self.text_password = wx.StaticText(self.m_panel1, wx.ID_ANY, u"����������", wx.DefaultPosition, wx.Size(150, 20),
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
                                           u"��������������Ҫץȡר����Ϣ�Ŀ�ʼ�ͽ������ڣ�\n���ڸ�ʽ20170731.��λ��������һ��Ҫ��ȫ0��", wx.DefaultPosition,
                                           wx.DefaultSize, 0)
        self.m_staticText5.Wrap(-1)
        self.m_staticText5.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText5.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer14.Add(self.m_staticText5, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer12.Add(bSizer14, 0, wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer16 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText6 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"��ʼ����:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText6.Wrap(-1)
        self.m_staticText6.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText6.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer16.Add(self.m_staticText6, 0, wx.ALL, 5)

        self.text_startdate = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          0)
        bSizer16.Add(self.text_startdate, 0, wx.ALL, 5)

        self.m_staticText7 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"��������:", wx.DefaultPosition, wx.DefaultSize, 0)
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

        self.button_exit = wx.Button(self, wx.ID_ANY, u"�˳�", wx.DefaultPosition, wx.Size(-1, 35), 0)
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
        self.updatedisplay("��ʼץȡ".decode('gbk'))
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
        #startdate_filter = startdate[0:4] + "%2F" + startdate[4:6] + "%2F" + startdate[6:8]
        #enddate_filter = enddate[0:4] + "%2F" + enddate[4:6] + "%2F" + enddate[6:8]

        # �ų������״̬
        list_except = ["����".decode('gbk'), '�˻ط�����'.decode('gbk'), '����'.decode('gbk')]
        # ģ���½
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
        # ��ȡ����
        # ��ʹ��limit=1����¼��ȡ���ֵ��
        url_data = "http://10.110.6.34/audit/bpm/complete"
        payload_1 = "limit=1&sortDirect=DESC&sortField=created_at"
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
        # ��ȡ���ֵ
        max_number = re.search(r'"pagination":{"currentPage":1,"offset":"1","total":(\d+),', data_1).groups()[0]
        #�޷�ֱ��ʹ�����ֵ��ȡ����ҳ��ȡȻ��ϲ�
        total_page = int(max_number) / 20 + 2
        print total_page
        # ��ҳ����ȡ��Ϣ
        list_status_temp = []
        list_date_created_temp = []
        list_creator_temp = []
        list_current_node_temp = []
        list_sn_filename_temp = []
        list_num_temp = []
        for page_number in range(1, total_page):
            print page_number
            payload_data = "limit=20&page={page}&sortDirect=DESC&sortField=created_at".format(page=page_number)
            #payload_data = "limit={max_number}&sortDirect=DESC&sortField=created_at".format(max_number=max_number)
            response_data = get_data.post(url_data, data=payload_data, headers=headers_data, verify=False)
            data_original = response_data.content
            #print data_original
            #��ȡ״̬
            list_status_temp_1 = re.findall(r'"status":"<span class=my_task_status_\w*?>(.*?)<\\/span>', data_original)
            list_status_temp_2 = [item.decode('unicode_escape') for item in list_status_temp_1]
            #print list_status_temp_2
            # ��ȡ����ʱ��
            list_date_created_temp_1 = re.findall(r'"created_at":"(\d+\\/\d+\\/\d+)', data_original)
            list_date_created_temp_2 = [item.replace("\\/", "-") for item in list_date_created_temp_1]
            #print list_date_created_temp_2
            #��ȡ������
            list_creator_temp_1 = re.findall(r',"created_by":"([^<].*?)",', data_original)
            list_creator_temp_3 = [item.decode('unicode_escape') for item in list_creator_temp_1]
            list_creator_temp_2 = [re.search(r"\D*", item).group() for item in list_creator_temp_3]
            #print list_creator_temp_1
            #��ȡ��ǰ�����ڵ�
            list_current_node_temp_1 = re.findall(r'"node_name":"<span class=\\"node-icon\\" data-bind=\w*? title=.*?><\\/span>(.*?)",', data_original)
            list_current_node_temp_2 = [item.decode('unicode_escape') for item in list_current_node_temp_1]
            #print  list_current_node_temp_2
            #��ȡ���Ӻͱ�ź��ļ���
            list_data_sn_filename_temp = re.findall(r'"subject":"<a href=\\"http:\\/\\/10.110.6.34\\/invention\\/inventions\\/view\\/(\d+)\\" target=\\"_blank\\">(.*?)<\\/a>",', data_original)
            list_num_temp_1 = []
            list_sn_filename_temp_1 = []
            for item in list_data_sn_filename_temp:
                list_num_temp_1.append(item[0])
                list_sn_filename_temp_1.append(item[1].decode('unicode_escape'))

            list_status_temp.extend(list_status_temp_2)
            list_creator_temp.extend(list_creator_temp_2)
            list_date_created_temp.extend(list_date_created_temp_2)
            list_current_node_temp.extend(list_current_node_temp_2)
            list_num_temp.extend(list_num_temp_1)
            list_sn_filename_temp.extend(list_sn_filename_temp_1)

        # �ȴ���һ�����ݣ���״̬���޳��б��ġ����ֱ�ɾ��ֻʣ��\/�ġ�SN�ظ���ȥ����׫д��Ϊ�����˳���Ϣ������������ȥ��
        list_status = []
        list_creator = []
        list_date_created = []
        list_current_node = []
        list_sn_filename = []
        list_num = []
        list_sn = []
        list_filename = []
        list_creator_except = ["�����˳���Ϣ��".decode('gbk'), "������".decode('gbk')]
        for index_status, item_status in enumerate(list_status_temp):
            if item_status not in list_except and list_sn_filename_temp[index_status] != '\/' and list_creator_temp[index_status] not in list_creator_except:
                if list_sn_filename_temp[index_status].split("\\/")[0] not in list_sn:
                    list_sn.append(list_sn_filename_temp[index_status].split("\\/")[0])
                    list_filename.append(list_sn_filename_temp[index_status].split("\\/")[1])
                    list_status.append(item_status)
                    list_creator.append(list_creator_temp[index_status])
                    list_date_created.append(list_date_created_temp[index_status])
                    list_current_node.append(list_current_node_temp[index_status])
                    list_sn_filename.append(list_sn_filename_temp[index_status])
                    list_num.append(list_num_temp[index_status])
        #��ȡ����
        list_link = ["http://10.110.6.34/invention/inventions/view/" + i for i in list_num ]
        #list_sn = [item.split("\\/")[0] for item in list_sn_filename]
        #list_filename = [item.split("\\/")[1] for item in list_sn_filename]
        print "sn length " + str(len(list_sn))
        print "status length " + str(len(list_status))
        print "current node length " + str(len(list_current_node))
        print "creator length " + str(len(list_creator))
        print "date created " + str(len(list_date_created))
        print  "filename lenngth " + str(len(list_filename))
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
        list_username_lastupdate = []
        list_type_invention = []
        list_date_lastupdate = []
        list_status_second = []
        list_department = []

        a = int(len(list_status) / 10)

        for index, item in enumerate(list_link):
            if index % a == 0:
                b = int(index / a) * 10
                self.updatedisplay(b)
            print item
            data_temp_temp = get_data.get(item, headers=headers_link, verify=False)
            if data_temp_temp.status_code != 404:
                data_temp = data_temp_temp.text
                data_soup_tobe_filter = BeautifulSoup(data_temp, "html.parser")
                # print data_soup_tobe_filter
                status_second = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(3) > td")[0].get_text().strip()
                type_invention = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(6) > td")[0].get_text().strip()
                department_temp = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(10) > td > a")
                department = "".join([i.get_text().strip() for i in department_temp])
                name_daili = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(14) > td > a")
                username_last_update_temp = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(21) > td")[0].get_text().strip().split( " ")[0]
                username_last_update = re.search(r"\D*", username_last_update_temp).group()
                date_last_update_temp = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(21) > td")[0].get_text().strip().split( " ")[1]
                date_last_update = date_last_update_temp.replace("/", "-")

                if len(name_daili) != 0:
                    list_data_daili.append(name_daili[0].get_text().strip())
                else:
                    list_data_daili.append("None")
                list_type_invention.append(type_invention)
                list_username_lastupdate.append(username_last_update)
                list_date_lastupdate.append(date_last_update)
                list_status_second.append(status_second)
                list_department.append(department)

        list_status_write = []
        list_sn_write = []
        list_filename_write = []
        list_department_write = []
        list_type_write = []
        list_creator_write = []
        list_date_created_write = []
        list_username_lastupdate_write = []
        list_date_lastupdate_write = []
        list_current_node_write = []
        list_name_daili_write = []

        list_status_second_except = ["׫д����".decode('gbk'),"������".decode('gbk')]
        for index_filter, item_filter in enumerate(list_status_second):
            if item_filter not in list_status_second_except:
                list_status_write.append(item_filter)
                list_sn_write.append(list_sn[index_filter])
                list_filename_write.append(list_filename[index_filter])
                list_department_write.append(list_department[index_filter])
                list_type_write.append(list_type_invention[index_filter])
                list_creator_write.append(list_creator[index_filter])
                list_date_created_write.append(list_date_created[index_filter])
                list_username_lastupdate_write.append(list_username_lastupdate[index_filter])
                list_date_lastupdate_write.append(list_date_lastupdate[index_filter])
                list_current_node_write.append(list_current_node[index_filter])
                list_name_daili_write.append(list_data_daili[index_filter])


        print "last sn length " + str(len(list_sn_write))
        print "last status length " + str(len(list_status_write))
        print "last current node length " + str(len(list_current_node_write))
        print "last creator length " + str(len(list_creator_write))
        print "last date created length" + str(len(list_date_created_write))
        print "last filename length " + str(len(list_filename_write))
        print "last username lastupdate length " + str(len(list_username_lastupdate_write))
        print "last date lastupdate length " + str(len(list_date_lastupdate_write))
        print "last department length " + str(len(list_department_write))
        print "last type length " + str(len(list_type_write))
        print "last daili length " + str(len(list_name_daili_write))

        title_sheet = ['��ǰ״̬'.decode('gbk'), '�᰸���'.decode('gbk'), '�᰸����'.decode('gbk'), '����'.decode('gbk'), 'ר������'.decode('gbk'), '׫д��'.decode('gbk'), '�ύʱ��'.decode('gbk'), '��������'.decode('gbk'), '������ʱ��'.decode('gbk'), '��ǰ�ڵ�'.decode('gbk'), '��������'.decode('gbk')]
        timestamp = time.strftime('%Y%m%d', time.localtime())
        # department_write = "������֤��".decode('gbk')
        workbook_display = xlsxwriter.Workbook('%sר������-%s.xlsx'.decode('gbk') % (department_write, timestamp))
        sheet = workbook_display.add_worksheet('2017����%sר������'.decode('gbk') % department_write)
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
        sheet.set_column('K:K', 33)
        sheet.merge_range(0, 0, 0, 10, "%s2017����ר������".decode('gbk') % department_write, formattitle)
        for index_title, item_title in enumerate(title_sheet):
            sheet.write(1, index_title, item_title, formatone)
        for index_data, item_data in enumerate(list_status_write):
            if item_data not in list_status_second_except:
                sheet.write(2 + index_data, 0, item_data, formatone)
                sheet.write(2 + index_data, 1, list_sn_write[index_data], formatone)
                sheet.write(2 + index_data, 2, list_filename_write[index_data], formatone)
                sheet.write(2 + index_data, 3, list_department_write[index_data], formatone)
                sheet.write(2 + index_data, 4, list_type_write[index_data], formatone)
                sheet.write(2 + index_data, 5, list_creator_write[index_data], formatone)
                sheet.write_datetime(2 + index_data, 6, datetime.datetime.strptime(list_date_created_write[index_data], '%Y-%m-%d'),
                                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))

                sheet.write(2 + index_data, 7, list_username_lastupdate_write[index_data], formatone)
                sheet.write_datetime(2 + index_data, 8, datetime.datetime.strptime(list_date_lastupdate_write[index_data], '%Y-%m-%d'),
                                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
                sheet.write(2 + index_data, 9, list_current_node_write[index_data], formatone)
                sheet.write(2 + index_data, 10, list_name_daili_write[index_data], formatone)
        workbook_display.close()
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        self.updatedisplay("ץȡ����,�����˳���ť�˳�����".decode('gbk'))
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
            self.output_info.AppendText("���".decode('gbk') + unicode(t) + "%")
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