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
import urllib2

print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
except_list = ["撰写驳回".decode('gbk'), '待决定'.decode('gbk')]
username = "yanshuo@inspur.com"
password = "sunyu1314ke"
startdate = "20170320"
enddate = "20171212"
startdate_filter = startdate[0:4] + "%2F" + startdate[4:6] + "%2F" + startdate[6:8]
enddate_filter = enddate[0:4] + "%2F" + enddate[4:6] + "%2F" + enddate[6:8]
# 模拟登陆
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
}
get_data = requests.session()
get_data.post(url_login, data=payload_login, headers=headers_base)

# 获取数据
# 先使用limit=1来登录获取最大值。
url_data = "http://10.110.6.34/invention/inventions/index"
payload_1 = "filter%5BInvention.created%5D%5Bfrom%5D={starttime}&filter%5BInvention.created%5D%5Bto%5D={endtime}&limit=1&sortDirect=&sortField=".format(starttime=startdate_filter, endtime=enddate_filter)
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
# 获取最大值
max_number = re.search(r'"pagination":{"currentPage":1,"offset":"1","total":(\d+),', data_1).groups()[0]

# 使用最大值来获取信息
# payload_data = "limit=%s" % max_number
payload_data = "limit={max_number}&filter%5BInvention.updated%5D%5Bfrom%5D={starttime}&filter%5BInvention.updated%5D%5Bto%5D={endtime}".format(
    max_number=max_number, starttime=startdate_filter, endtime=enddate_filter)
response_data = get_data.post(url_data, data=payload_data, headers=headers_data, verify=False)
data_original = response_data.content
# 获取编号
list_data_sn = re.findall(r'\"Invention.track_number\"\:\"(\d+)"', data_original)
# 获取链接
# 获取链接的数字
data_link_temp = re.findall(r'"Invention.title":"<a href=\\"http:\\/\\/10.110.6.34\\/invention\\/inventions\\/view\\/(\d+)\\" target=\\"_blank\\"', data_original)
# 再将数字连接到前置地址上，形成一级地址
list_data_link = ["http://10.110.6.34/invention/inventions/view/" + i for i in data_link_temp]
# 将数字连接获取审批日志地址
list_data_link_log = ["http://10.110.6.34/invention/inventions/audit_logs/" + i for i in data_link_temp]
# 获取专利名称。先获取返回值，然后再转换编码
data_name_temp = re.findall(r'"Invention.title":"<a.*?target=\\"_blank\\">(.*?)<\\/a>', data_original)
list_data_name = [i.decode('unicode_escape') for i in data_name_temp]
# 获取部门和处。先获取返回值，然后再处理编码和替换多余字符
data_department_temp = re.findall(r'"Invention.organization":"<a.*?title=(.*?)>', data_original)
list_data_department = [i.decode('unicode_escape').replace(" &gt; ", "") for i in data_department_temp]
# 获取创建时间。先获取返回值，然后替换字符
data_created_date_temp = re.findall(r'"Invention.created":"(\d+\\/\d+\\/\d+)"', data_original)
list_data_created_date = [i.replace("\\/", "-") for i in data_created_date_temp]
# 获取更新时间。先获取返回值，然后替换字符
data_update_date_temp = re.findall(r'"Invention.updated":"(\d+\\/\d+\\/\d+)"', data_original)
list_data_update_date = [i.replace("\\/", "-") for i in data_update_date_temp]
# 获取当前状态.先获取返回值，然后再转换编码
data_status_temp = re.findall(r'"Invention.node_status":"<a href=.*?>(.*?)<\\/a>', data_original)
list_data_status = [i.decode('unicode_escape') for i in data_status_temp]
# 先处理一遍数据，把撰写驳回或者加上待决定的去除
list_status = []
list_sn = []
list_link_one = []
list_link_log = []
list_name = []
list_department = []
list_created_date = []
list_updated_date = []
for index_status, item_status in enumerate(list_data_status):
    if item_status not in except_list:
        list_status.append(item_status)
        list_sn.append(list_data_sn[index_status])
        list_link_one.append(list_data_link[index_status])
        list_link_log.append(list_data_link_log[index_status])
        list_name.append(list_data_name[index_status])
        list_department.append(list_data_department[index_status])
        list_created_date.append(list_data_created_date[index_status])
        list_updated_date.append(list_data_update_date[index_status])
print "node1 " + str(len(list_updated_date))
headers_link_one = {
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
#开始获取每个专利的首页数据
list_data_daili = []
list_data_name_lastupdate = []
list_type_invention = []
list_username_created = []

a = int(len(list_status) / 10)
# print "a=%s" % str(a)
for index, item in enumerate(list_link_one):
    if index % a == 0:
        b = int(index / a) * 10
        print "Progess %s" % str(b) + "%"

    data_temp = get_data.get(item, headers=headers_link_one, verify=False).text
    data_soup_tobe_filter = BeautifulSoup(data_temp, "html.parser")
    type_invention = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(6) > td")[0].get_text().strip()
    name_daili = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(14) > td > a")
    name_last_update_temp = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(21) > td")[0].get_text().strip().split(" ")[0]
    name_last_update = re.search(r"\D*", name_last_update_temp).group()
    name_creator_temp = data_soup_tobe_filter.select(".major-left > div > table > tr:nth-of-type(20) > td")[0].get_text().strip().split(" ")[0]
    name_creator = re.search(r"\D*", name_creator_temp).group()
    # print "creator %s" % name_creator
    if len(name_daili) != 0:
        list_data_daili.append(name_daili[0].get_text().strip())
    else:
        list_data_daili.append("None")
    list_type_invention.append(type_invention)
    list_data_name_lastupdate.append(name_last_update)
    list_username_created.append(name_creator)


#开始获取每个专利的审批日志数据
headers_get_log = {
    'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
    'accept-encoding': "gzip, deflate",
    'accept-language': "zh-CN,zh;q=0.8",
    'cache-control': "no-cache",
    'connection': "keep-alive",
    'host': "10.110.6.34",
    'upgrade-insecure-requests': "1",
    'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
    }
list_tijiao_date_start = []
list_tijiao_date_jiekouren = []
list_tijiao_date_lingdao = []
list_tijiao_date_final = []
list_zhuanxie_date_start = []
list_zhuanxie_date_daili = []
list_zhuanxie_date_creator_confirm = []
list_zhuanxie_date_final = []
list_except_process = ["驳回".decode('gbk'), "撤回".decode('gbk'), "退回".decode('gbk')]
for index_log, item_log in enumerate(list_link_log):
    tijiao_date_final = "None"
    tijiao_date_lingdao = "None"
    tijiao_date_jiekouren = "None"
    tijiao_date_start = "None"
    zhuanxie_date_final = "None"
    zhuanxie_date_creator_confirm = "None"
    zhuanxie_date_daili = "None"
    zhuanxie_date_start = "None"
    data_temp_log = get_data.get(item_log, headers=headers_get_log, verify=False).text
    data_log_tobe_filter = BeautifulSoup(data_temp_log, "html.parser")
    # 获取状态，是提案中还是撰写中。等等
    status_second = data_log_tobe_filter.select(".major-title > span")[0].get_text().strip()
    #获取流程名字
    list_audit_name_temp_1 = re.findall(r'"audit_name":"((?:\\\w+){0,20})",', data_temp_log)
    list_audit_name_temp_2 = [i.decode('unicode_escape') for i in list_audit_name_temp_1]
    #获取log id
    list_logid_temp = re.findall(r'"run_id":"(\w+)",', data_temp_log)
    #根据 logid生成log的url
    list_logid_url_temp_1 = ["http://10.110.6.34/audit/bpm/async_logs/" + i for i in list_logid_temp]
    #分别储存撰写和提交流程。
    list_zhuanxie_process_url_temp_1 = []
    list_tijiao_process_url_temp_1 = []

    for index_audit, item_audit in enumerate(list_audit_name_temp_2):
        if item_audit == "提交代理所撰写流程".decode('gbk'):
            list_zhuanxie_process_url_temp_1.append(list_logid_url_temp_1[index_audit])
        if item_audit == "浪潮信息交底书审核流程".decode('gbk'):
            list_tijiao_process_url_temp_1.append(list_logid_url_temp_1[index_audit])

    #根据当前状态开始分情况处理
    list_status_1 = ["提案中".decode('gbk'), "提案通过".decode('gbk')]
    list_status_2 = ["撰写中".decode('gbk'), "撰写通过".decode('gbk'), "申请专利".decode('gbk')]
    #先处理在提案流程中的
    if status_second in list_status_1:
        for index_tijiao, item_tijiao in enumerate(list_tijiao_process_url_temp_1):
            data_temp_tijiao = get_data.get(item_tijiao, headers=headers_get_log, verify=False).text
            #获取审批结果
            list_tijiao_action_temp_1 = re.findall(r',"action_name":"((?:\\\w+){0,20})","', data_temp_tijiao)
            list_tijiao_action_temp_2 = [i.decode('unicode_escape') for i in list_tijiao_action_temp_1]
            if list_tijiao_action_temp_2[0] in list_except_process:
                continue
            else:
                # 节点名称
                list_tijiao_activity_name_temp_1 = re.findall(r'{"activity_name":"((?:\\\w+){0,20})"', data_temp_tijiao)
                list_tijiao_activity_name_temp_3 = [i.decode('unicode_escape') for i in list_tijiao_activity_name_temp_1]
                list_tijiao_activity_name_temp_2 = []
                # 节点时间
                list_tijiao_created_at_temp_1 = re.findall(r'"created_at":"(\d+\\/\d+\\/\d+)', data_temp_tijiao)
                list_tijiao_created_at_temp_3 = [i.replace("\\/", "-") for i in list_tijiao_created_at_temp_1]
                list_tijiao_created_at_temp_2 = []
                # 处理节点数据，排除掉重复的节点，如专利工程师审核阶段进行转交等情况。
                for index_tijiao_actitity_name, item_tijiao_activity_name in enumerate(list_tijiao_activity_name_temp_3):
                    if item_tijiao_activity_name not in list_tijiao_activity_name_temp_2:
                        list_tijiao_activity_name_temp_2.append(item_tijiao_activity_name)
                        list_tijiao_created_at_temp_2.append(list_tijiao_created_at_temp_3[index_tijiao_actitity_name])

                if list_tijiao_activity_name_temp_2[0] == "专利工程师审核".decode('gbk'):
                    tijiao_date_final = list_tijiao_created_at_temp_2[0]
                    tijiao_date_lingdao = list_tijiao_created_at_temp_2[1]
                    tijiao_date_jiekouren = list_tijiao_created_at_temp_2[2]
                    tijiao_date_start = list_tijiao_created_at_temp_2[3]
                elif list_tijiao_activity_name_temp_2[0] == "部门领导审核".decode('gbk'):
                    tijiao_date_final = "None"
                    tijiao_date_lingdao = list_tijiao_created_at_temp_2[0]
                    tijiao_date_jiekouren = list_tijiao_created_at_temp_2[1]
                    tijiao_date_start = list_tijiao_created_at_temp_2[2]
                elif list_tijiao_activity_name_temp_2[0] == "接口人审核".decode('gbk'):
                    tijiao_date_final = "None"
                    tijiao_date_lingdao = "None"
                    tijiao_date_jiekouren = list_tijiao_created_at_temp_2[0]
                    tijiao_date_start = list_tijiao_created_at_temp_2[1]
                elif list_tijiao_activity_name_temp_2[0] == "开始节点".decode('gbk'):
                    tijiao_date_final = "None"
                    tijiao_date_lingdao = "None"
                    tijiao_date_jiekouren = "None"
                    tijiao_date_start = list_tijiao_created_at_temp_2[0]
                else:
                    tijiao_date_final = "None"
                    tijiao_date_lingdao = "None"
                    tijiao_date_jiekouren = "None"
                    tijiao_date_start = "None"
                break

        list_tijiao_date_final.append(tijiao_date_final)
        list_tijiao_date_lingdao.append(tijiao_date_lingdao)
        list_tijiao_date_jiekouren.append(tijiao_date_jiekouren)
        list_tijiao_date_start.append(tijiao_date_start)

        list_zhuanxie_date_final.append("None")
        list_zhuanxie_date_creator_confirm.append("None")
        list_zhuanxie_date_daili.append("None")
        list_zhuanxie_date_start.append("None")
    # 再处理在撰写流程中的
    elif status_second in list_status_2:
        for index_zhuanxie, item_zhuanxie in enumerate(list_zhuanxie_process_url_temp_1):
            data_temp_zhuanxie = get_data.get(item_zhuanxie, headers=headers_get_log, verify=False).text
            # 获取审批结果
            list_zhuanxie_action_temp_1 = re.findall(r',"action_name":"((?:\\\w+){0,20})","', data_temp_zhuanxie)
            list_zhuanxie_action_temp_2 = [i.decode('unicode_escape') for i in list_zhuanxie_action_temp_1]
            if list_zhuanxie_action_temp_2[0] in list_except_process:
                continue
            else:
                #节点名称
                list_zhuanxie_activity_name_temp_1 = re.findall(r'{"activity_name":"((?:\\\w+){0,20})"', data_temp_zhuanxie)
                list_zhuanxie_activity_name_temp_3 = [i.decode('unicode_escape') for i in list_zhuanxie_activity_name_temp_1]
                list_zhuanxie_activity_name_temp_2 = []
                #节点时间
                list_zhuanxie_created_at_temp_1 = re.findall(r'"created_at":"(\d+\\/\d+\\/\d+)', data_temp_zhuanxie)
                list_zhuanxie_created_at_temp_3 = [i.replace("\\/", "-") for i in list_zhuanxie_created_at_temp_1]
                list_zhuanxie_created_at_temp_2 = []
                # 处理节点数据，排除掉重复的节点，如专利工程师审核阶段进行转交等情况。
                for index_zhuanxie_actitity_name, item_zhuanxie_activity_name in enumerate(list_zhuanxie_activity_name_temp_3):
                    if item_zhuanxie_activity_name not in list_zhuanxie_activity_name_temp_2:
                        list_zhuanxie_activity_name_temp_2.append(item_zhuanxie_activity_name)
                        list_zhuanxie_created_at_temp_2.append(list_zhuanxie_created_at_temp_3[index_zhuanxie_actitity_name])

                if list_zhuanxie_activity_name_temp_2[0] == "专利工程师确认".decode('gbk'):
                    zhuanxie_date_final = list_zhuanxie_created_at_temp_2[0]
                    zhuanxie_date_creator_confirm = list_zhuanxie_created_at_temp_2[1]
                    zhuanxie_date_daili = list_zhuanxie_created_at_temp_2[2]
                    zhuanxie_date_start = list_zhuanxie_created_at_temp_2[3]
                elif list_zhuanxie_activity_name_temp_2[0] == "发明人确认".decode('gbk'):
                    zhuanxie_date_final = "None"
                    zhuanxie_date_creator_confirm = list_zhuanxie_created_at_temp_2[0]
                    zhuanxie_date_daili = list_zhuanxie_created_at_temp_2[1]
                    zhuanxie_date_start = list_zhuanxie_created_at_temp_2[2]
                elif list_zhuanxie_activity_name_temp_2[0] == "代理人撰写稿上传".decode('gbk'):
                    zhuanxie_date_final = "None"
                    zhuanxie_date_creator_confirm = "None"
                    zhuanxie_date_daili = list_zhuanxie_created_at_temp_2[0]
                    zhuanxie_date_start = list_zhuanxie_created_at_temp_2[1]
                elif list_zhuanxie_activity_name_temp_2[0] == "开始节点".decode('gbk'):
                    zhuanxie_date_final = "None"
                    zhuanxie_date_creator_confirm = "None"
                    zhuanxie_date_daili = "None"
                    zhuanxie_date_start = list_zhuanxie_created_at_temp_2[0]
                else:
                    zhuanxie_date_final = "None"
                    zhuanxie_date_creator_confirm = "None"
                    zhuanxie_date_daili = "None"
                    zhuanxie_date_start = "None"
                break

        list_zhuanxie_date_final.append(zhuanxie_date_final)
        list_zhuanxie_date_creator_confirm.append(zhuanxie_date_creator_confirm)
        list_zhuanxie_date_daili.append(zhuanxie_date_daili)
        list_zhuanxie_date_start.append(zhuanxie_date_start)

        #撰写流程肯定要伴随一个前置提交流程。但是也会有异常情况，发明人直接提交撰写。需要排除。
        #如果只有撰写，但是没有提交流程。提交信息全部是None.
        if len(list_tijiao_process_url_temp_1) == 0:
            pass
        else:
            for index_tijiao, item_tijiao in enumerate(list_tijiao_process_url_temp_1):
                data_temp_tijiao = get_data.get(item_tijiao, headers=headers_get_log, verify=False).text
                # 获取审批结果
                list_tijiao_action_temp_1 = re.findall(r',"action_name":"((?:\\\w+){0,20})","', data_temp_tijiao)
                list_tijiao_action_temp_2 = [i.decode('unicode_escape') for i in list_tijiao_action_temp_1]
                if list_tijiao_action_temp_2[0] in list_except_process:
                    #如果有撰写流程，但是提交流程却被博汇。则说明此条专利流程也是错误的。时间点全部都是None.
                    continue
                else:
                    # 节点名称
                    list_tijiao_activity_name_temp_1 = re.findall(r'{"activity_name":"((?:\\\w+){0,20})"', data_temp_tijiao)
                    list_tijiao_activity_name_temp_3 = [i.decode('unicode_escape') for i in list_tijiao_activity_name_temp_1]
                    list_tijiao_activity_name_temp_2 = []
                    # 节点时间
                    list_tijiao_created_at_temp_1 = re.findall(r'"created_at":"(\d+\\/\d+\\/\d+)', data_temp_tijiao)
                    list_tijiao_created_at_temp_3 = [i.replace("\\/", "-") for i in list_tijiao_created_at_temp_1]
                    list_tijiao_created_at_temp_2 = []
                    #处理节点数据，排除掉重复的节点，如专利工程师审核阶段进行转交等情况。
                    for index_tijiao_actitity_name, item_tijiao_activity_name in enumerate(list_tijiao_activity_name_temp_3):
                        if item_tijiao_activity_name not in list_tijiao_activity_name_temp_2:
                            list_tijiao_activity_name_temp_2.append(item_tijiao_activity_name)
                            list_tijiao_created_at_temp_2.append(list_tijiao_created_at_temp_3[index_tijiao_actitity_name])

                    if list_tijiao_activity_name_temp_2[0] == "专利工程师审核".decode('gbk'):
                        tijiao_date_final = list_tijiao_created_at_temp_2[0]
                        tijiao_date_lingdao = list_tijiao_created_at_temp_2[1]
                        tijiao_date_jiekouren = list_tijiao_created_at_temp_2[2]
                        tijiao_date_start = list_tijiao_created_at_temp_2[3]
                    elif list_tijiao_activity_name_temp_2[0] == "部门领导审核".decode('gbk'):
                        tijiao_date_final = "None"
                        tijiao_date_lingdao = list_tijiao_created_at_temp_2[0]
                        tijiao_date_jiekouren = list_tijiao_created_at_temp_2[1]
                        tijiao_date_start = list_tijiao_created_at_temp_2[2]
                    elif list_tijiao_activity_name_temp_2[0] == "接口人审核".decode('gbk'):
                        tijiao_date_final = "None"
                        tijiao_date_lingdao = "None"
                        tijiao_date_jiekouren = list_tijiao_created_at_temp_2[0]
                        tijiao_date_start = list_tijiao_created_at_temp_2[1]
                    elif list_tijiao_activity_name_temp_2[0] == "开始节点".decode('gbk'):
                        tijiao_date_final = "None"
                        tijiao_date_lingdao = "None"
                        tijiao_date_jiekouren = "None"
                        tijiao_date_start = list_tijiao_created_at_temp_2[0]
                    else:
                        tijiao_date_final = "None"
                        tijiao_date_lingdao = "None"
                        tijiao_date_jiekouren = "None"
                        tijiao_date_start = "None"
                    #获取到一次没被驳回的提交流程结果就退出整个for循环。
                    break

        list_tijiao_date_final.append(tijiao_date_final)
        list_tijiao_date_lingdao.append(tijiao_date_lingdao)
        list_tijiao_date_jiekouren.append(tijiao_date_jiekouren)
        list_tijiao_date_start.append(tijiao_date_start)
    print len(list_tijiao_date_final)

#写入xlsx文件
title_sheet = ['当前状态'.decode('gbk'), '提案编号'.decode('gbk'), '提案名称'.decode('gbk'), '处别'.decode('gbk'),
               '专利类型'.decode('gbk'), '撰写人'.decode('gbk'), '创建时间'.decode('gbk'),
               '最后更新人'.decode('gbk'), '最后更新时间'.decode('gbk'), '代理名称'.decode('gbk'), '发明人提起流程时间'.decode('gbk'), '接口人审批时间'.decode('gbk'),
               '领导审批时间'.decode('gbk'), '专利工程师确认时间'.decode('gbk'), '撰写开始时间'.decode('gbk'), '代理提交撰写稿时间'.decode('gbk'),
               '发明人确认撰写稿时间'.decode('gbk'), '专利工程师最终确认时间'.decode('gbk')]
timestamp = time.strftime('%Y%m%d', time.localtime())
department_write = "测试验证部".decode('gbk')
workbook_display = xlsxwriter.Workbook('%s专利总览包含各节点时间-%s.xlsx'.decode('gbk') % (department_write, timestamp))
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
sheet.set_column('K:R', 14)

sheet.merge_range(0, 0, 0, 17, "%s2017财年专利总览".decode('gbk') % department_write, formattitle)
for index_title, item_title in enumerate(title_sheet):
    sheet.write(1, index_title, item_title, formatone)
for index_data, item_data in enumerate(list_status):
    sheet.write(2 + index_data, 0, item_data, formatone)
    sheet.write(2 + index_data, 1, list_sn[index_data], formatone)
    sheet.write(2 + index_data, 2, list_name[index_data], formatone)
    sheet.write(2 + index_data, 3, list_department[index_data], formatone)
    sheet.write(2 + index_data, 4, list_type_invention[index_data], formatone)
    sheet.write(2 + index_data, 5, list_username_created[index_data], formatone)
    sheet.write_datetime(2 + index_data, 6, datetime.datetime.strptime(list_created_date[index_data], '%Y-%m-%d'),
                         workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
    sheet.write(2 + index_data, 7, list_data_name_lastupdate[index_data], formatone)
    sheet.write_datetime(2 + index_data, 8, datetime.datetime.strptime(list_data_update_date[index_data], '%Y-%m-%d'),
                         workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
    sheet.write(2 + index_data, 9, list_data_daili[index_data], formatone)

    #if list_tijiao_date_start[index_data] == "None":
    sheet.write(2 + index_data, 10, list_tijiao_date_start[index_data], formatone)
    #else:
    #    sheet.write_datetime(2 + index_data, 10, datetime.datetime.strptime(list_tijiao_date_start[index_data], '%Y-%m-%d'),
    #                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))

    #if list_tijiao_date_jiekouren[index_data] == "None":
    sheet.write(2 + index_data, 11, list_tijiao_date_jiekouren[index_data], formatone)
    #else:
    #    sheet.write_datetime(2 + index_data, 11, datetime.datetime.strptime(list_tijiao_date_jiekouren[index_data], '%Y-%m-%d'),
    #                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))

    #if list_tijiao_date_lingdao[index_data] == "None":
    sheet.write(2 + index_data, 12, list_tijiao_date_lingdao[index_data], formatone)
    #else:
    #    sheet.write_datetime(2 + index_data, 12, datetime.datetime.strptime(list_tijiao_date_lingdao[index_data], '%Y-%m-%d'),
    #                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))

    #if list_tijiao_date_final[index_data] == "None":
    sheet.write(2 + index_data, 13, list_tijiao_date_final[index_data], formatone)
    #else:
    #    sheet.write_datetime(2 + index_data, 13, datetime.datetime.strptime(list_tijiao_date_final[index_data], '%Y-%m-%d'),
    #                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))

    #if list_zhuanxie_date_start[index_data] == "None":
    sheet.write(2 + index_data, 14, list_zhuanxie_date_start[index_data], formatone)
    #else:
    #    sheet.write_datetime(2 + index_data, 14, datetime.datetime.strptime(list_zhuanxie_date_start[index_data], '%Y-%m-%d'),
    #                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))

    #if list_zhuanxie_date_daili[index_data] == "None":
    sheet.write(2 + index_data, 15, list_zhuanxie_date_daili[index_data], formatone)
    #else:
    #    sheet.write_datetime(2 + index_data, 15, datetime.datetime.strptime(list_zhuanxie_date_daili[index_data], '%Y-%m-%d'),
    #                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))

    #if list_zhuanxie_date_creator_confirm[index_data] == "None":
    sheet.write(2 + index_data, 16, list_zhuanxie_date_creator_confirm[index_data], formatone)
    #else:
    #    sheet.write_datetime(2 + index_data, 16, datetime.datetime.strptime(list_zhuanxie_date_creator_confirm[index_data], '%Y-%m-%d'),
    #                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))

    #if list_zhuanxie_date_final[index_data] == "None":
    sheet.write(2 + index_data, 17, list_zhuanxie_date_final[index_data], formatone)
    #else:
    #    sheet.write_datetime(2 + index_data, 17, datetime.datetime.strptime(list_zhuanxie_date_final[index_data], '%Y-%m-%d'),
    #                     workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))

workbook_display.close()
print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
