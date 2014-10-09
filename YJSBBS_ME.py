# -*- coding: utf8 -*-
#程序：应届生BBS爬虫（上海机械招聘）
#版本：0.1
#作者：wwb
#日期：2014-01-19
#语言：Python 2.7.6
#平台：windows系统
#操作：输入带分页的地址，去掉最后面的数字，设置一下起始页数和终点页数。网站url的参数部分把页码直接挂上去了，可以用一下该特点实现翻页
#功能：下载过去几天内有效页码内的所有招聘信息（公司职位 发布日期）存放在excel中（链接能用）
#地址：http://www.yingjiesheng.com/major/jixie/shanghai/list_+p+.html
#页面特点：<h3><span></span><a></a></h3>
#------------------代码规划：-----------------------
# 1）输入天数,输出这几天的招聘信息（格式暂定excel），正则表达式的写法
# 2) 循环执行，抓取页面直到条件不满足（日期超过输入天数范围）
#    一种是正则表达式的筛选写法；一种是利用lxml解析xml文件
# 3）处理页面中的有效信息 存入指定文件中（文件名就是当天日期，如2014-01-16）正则表达式的写法
#------------------ToDo List：-----------------------
# 1）扩展应届生网更多行业招聘信息；
# 2) 扩展更多招聘网站信息；
# 3）提高抓取效率；直接打开网页文件分析内容与先下载到本地再打开文件分析，多线程？
# 4）是否被重复抓取过，用MongoDB记录一下已有信息如何？
# 5) 支持更多条件筛选，时间向前向后，某个时间段内

import codecs
import sys
from lxml import etree
import string
import urllib2
from pyExcelerator import *
import datetime

def JX_SH(dayCount):
    #excel格式设置
    fnt = Font()  
    fnt.name = '宋体'  
    fnt.colour_index = 4  
    fnt.bold = True  
    fnt.weight = 20
    
    borders = Borders()  
    borders.left = 6  
    borders.right = 6  
    borders.top = 6  
    borders.bottom = 6
    borders.width = '100'
      
    al = Alignment()  
    al.horz = Alignment.HORZ_CENTER  
    al.vert = Alignment.VERT_CENTER  
      
    style = XFStyle()
    style.font = fnt
    style.borders = borders
    style.alignment = al

    #生成新的excel
    w = Workbook()
    ws = w.add_sheet('ShangHai')
    k = 0;#标记是否有新的招聘信息
    today = datetime.datetime.now()
    for i in range(1,20):
        url = 'http://www.yingjiesheng.com/major/jixie/shanghai/list_'+str(i)+'.html';        
        html = urllib2.urlopen(url).read().decode('gbk')
        tree = etree.HTML(html)
        dates = tree.xpath("//h3//span[@class='date']")
        jobs = tree.xpath("//h3//a")
        length = len(dates)

        #进行填充
        j = 0;        
        for i in range(length):
            date1 = dates[i].text
            date2 = datetime.datetime.strptime(date1,"%Y-%m-%d")
            if((today-date2).days <= dayCount):                
                ws.write(k,0,k+1,style)
                ws.write(k,1,date1,style)
                ws.write(k,2,jobs[i].text,style)
                ws.write(k,3,jobs[i].get('href'),style)
                k=k+1
            else:
                j = 1
                break
        if(j==1):
            break
    if(k!=0):
        #不支持xlsx格式 所以用了xls
        w.save('Shanghai ME'+today.strftime("%Y%m%d%S")+'.xls')
        print '抓取招聘信息成功！'
    else:
        print '没有招聘信息！'
#if __name__ == '__main__': 转成exe不工作 取消掉
count = int(input(u'请输入要查询的日期范围(查询当天信息就输入1，最近一周信息就输入7)：\n')) 
JX_SH(count)
    

