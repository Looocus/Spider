# Spider
Get the lessons score from the university's official website.
#1.Get the score information---------CN
#2.Create a Excel file to save the information
#-*- coding:utf8 -*-
#__author__ = 'Danny'

import requests
import urllib
from lxml import etree
import re
import time
import http.cookiejar
import getpass
from lxml import etree
from lxml import html
import requests
import xlwt
import getpass

def getscorescr_std(username,password):
    ts_url = 'http://elearning.ustb.edu.cn/choose_courses/j_spring_security_check'
    headers = {}
    headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:55.0) Gecko/20100101 Firefox/55.0'
    postdata = {}
    postdata['j_username'] = username+',undergraduate'
    postdata['j_password'] = password
    postdata = urllib.parse.urlencode(postdata).encode('utf-8')
    #cookie
    cookie = http.cookiejar.MozillaCookieJar()
    handler = urllib.request.HTTPCookieProcessor(cookie)
    opener = urllib.request.build_opener(handler)
    req = urllib.request.Request(ts_url,postdata,headers)
    res = opener.open(req)
    req_1 = urllib.request.Request('http://elearning.ustb.edu.cn/choose_courses/loginsucc.action')
    res_1 = opener.open(req_1)
    req_2 = requests.get("http://elearning.ustb.edu.cn/choose_courses/information/singleStuInfo_singleStuInfo_loadSingleStuScorePage.action",cookies=cookie)
    req_2.encoding = 'utf-8'
    req_html = etree.HTML(req_2.content)
    items = req_html.xpath('//table/tbody/tr')
    #print(req_html)
    #print(items)
    return req_html,items
#获取学期
def getdate(requestss,itemss):
    Date = []
    for num in range(len(itemss)):
        path = '//table/tbody/tr[%d]/td[1]/text()' % (num+1)
        Date.append(requestss.xpath(path)[0])
    return Date
#写TXT,//*[@id="w_stuscore-body"]/table/tbody/tr[1]/td[1]
#获取课程编号
def getclnum(requestss,itemss):
    Clnum = []
    for num in range(len(itemss)):
        path = '//table/tbody/tr[%d]/td[2]/text()' % (num+1)
        Clnum.append(requestss.xpath(path)[0])
    return Clnum
#获取课程名称
def getclnme(requestss,itemss):
    Clnme = []
    for num in range(len(itemss)):
        path = '//table/tbody/tr[%d]/td[3]/text()' % (num+1)
        Clnme.append(requestss.xpath(path)[0])
    return Clnme
#获取课程类别
def getclseg(requestss,itemss):
    Clseg = []
    for num in range(len(itemss)):
        path = '//table/tbody/tr[%d]/td[4]/text()' % (num+1)
        Clseg.append(requestss.xpath(path)[0])
    return Clseg
#获取课程课时//*[@id="w_stuscore-body"]/table/tbody/tr[1]/td[6]
def getclhor(requestss,itemss):
    Clhor = []
    for num in range(len(itemss)):
        path = '//table/tbody/tr[%d]/td[6]/text()' % (num+1)
        Clhor.append(requestss.xpath(path)[0])
    return Clhor
#获取课程成绩
def getclsco(requestss,itemss):
    Clsco = []
    for num in range(len(itemss)):
        path = '//table/tbody/tr[%d]/td[7]/text()' % (num+1)
        Clsco.append(requestss.xpath(path)[0])
    return Clsco
#写入Excel
def WriteExcel(Date,Clnum,Clnme,Clseg,Clhor,Clsco):
    Score_wbk = xlwt.Workbook()
    sheet1 = Score_wbk.add_sheet('sheet 1')
    sheet1.write(0, 0, '课程学期')
    sheet1.write(0, 1, '课程编号')
    sheet1.write(0, 2, '课程名称')
    sheet1.write(0, 3, '课程类别')
    sheet1.write(0, 4, '课程课时')
    sheet1.write(0, 5, '课程成绩')
###############
    num_Date = 1
    for item in Date:
        sheet1.write(num_Date, 0, item)
        num_Date += 1
    num_Clnum = 1
    for item in Clnum:
        sheet1.write(num_Clnum, 1, item)
        num_Clnum += 1
    num_Clnme = 1
    for item in Clnme:
        sheet1.write(num_Clnme, 2, item)
        num_Clnme += 1
    num_Clseg = 1
    for item in Clseg:
        sheet1.write(num_Clseg, 3, item)
        num_Clseg += 1
    num_Clhor = 1
    for item in Clhor:
        sheet1.write(num_Clhor, 4, item)
        num_Clhor += 1
    num_Clsco = 1
    for item in Clsco:
        sheet1.write(num_Clsco, 5, item)
        num_Clsco += 1
    Score_wbk.save(u'ScoreUstb.xls')
#####################################
if __name__ == "__main__":
    username = input("学号：")
    password = input("密码：")
    getscorescr_std(username,password)
    requestss,itemss = getscorescr_std(username,password)
    Date = getdate(requestss,itemss)
    Clnum = getclnum(requestss,itemss)
    Clnme = getclnme(requestss,itemss)
    Clseg = getclseg(requestss,itemss)
    Clhor = getclhor(requestss,itemss)
    Clsco = getclsco(requestss,itemss)
    WriteExcel(Date, Clnum, Clnme, Clseg, Clhor, Clsco)
