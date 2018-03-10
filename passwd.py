# -*- coding=utf-8 -*-
import xlrd
#导入xlrd模块
import xlwt
import urllib
import urllib2
import cookielib
import os
from threading import Timer
from xlwt import Workbook, Formula
from xlrd import open_workbook
from xlutils.copy import copy

#打开指定文件路径的excel文件
xlsfile = r'passwd.xls'
passwd = xlrd.open_workbook(xlsfile)     #获得excel的book对象
#获得指定索引的sheet名字
sheet_name=passwd.sheet_names()[0]
#通过sheet名字来获取，如果知道sheet名字了可以直接指定
sheet=passwd.sheet_by_name(sheet_name)
#sheet0=passwd.sheet_by_index(0)     #通过sheet索引获得sheet对象
#passGenRuleRow = []
def passwd(k):
    if sheet.cell(k,1).value == 1.0:
        bigLetter = 'true'
    else:
        bigLetter = 'false'

    if sheet.cell(k,2).value == 1.0:
        lowerLetter = 'true'
    else:
        lowerLetter = 'false'

    if sheet.cell(k,3).value == 1.0:
        num = 'true'
    else:
        num = 'false'

    if sheet.cell(k,4).value == 1.0:
        specSyml = 'true'
    else:
        specSyml = 'false'
    while bigLetter =="false" and  lowerLetter == "false" and num == "false" and specSyml == "false":
        k+=1

        if sheet.cell(k, 1).value == 1.0:
            bigLetter = 'true'
        else:
            bigLetter = 'false'

        if sheet.cell(k, 2).value == 1.0:
            lowerLetter = 'true'
        else:
            lowerLetter = 'false'

        if sheet.cell(k, 3).value == 1.0:
            num = 'true'
        else:
            num = 'false'

        if sheet.cell(k, 4).value == 1.0:
            specSyml = 'true'
        else:
            specSyml = 'false'

    length = str(int(sheet.cell(k,5).value))
    requestPasswdAlogAddrs = "https://suijimimashengcheng.51240.com/web_system/51240_com_www/system/file/suijimimashengcheng/get/?dx=" + bigLetter + "&xx=" + lowerLetter + "&sz=" + num + "&fh=" + specSyml +"&fh_value=!%40%23%24%2525%255E%26*&cd=" + length
    up = urllib.urlopen(requestPasswdAlogAddrs).read()
    a=up.find("value");
    b=up.find("\" readonly")
    result=up[a+7:b]
    #wb.save('C:\Users\Administrator\Desktop\passwd.xlsx')
    #open_workbook(filepath)
    rb = open_workbook('D:\\python\\password\\passwd.xls')
    # 通过sheet_by_index()获取的sheet没有write()方法
    rs = rb.sheet_by_index(0)
    wb = copy(rb)
    # 通过get_sheet()获取的sheet有write()方法
    ws = wb.get_sheet(0)
    ws.write(k, 6, result)
    print k
    #save(filepath)
    wb.save('D:\\python\\password\\passwd.xls')

for i in range(1,1025):
    passwd(i)