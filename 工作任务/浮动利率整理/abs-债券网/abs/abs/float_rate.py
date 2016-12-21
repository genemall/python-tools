#encoding:utf-8
#解决中文乱码问题
import sys
reload(sys)
sys.setdefaultencoding( "utf-8")

import xlrd
import xlwt
import re
import urllib2
#from bs4 import BeautifulSoup

#下载网页成 html语言
def download(url, user_agent='wswp', num_retries=2):
    #print ('downloading:', url)
    headers = {'User-agent': user_agent}
    request = urllib2.Request(url, headers=headers)
    try:
        html = urllib2.urlopen(request).read()
    except urllib2.URLError as e:
        print ('Download error:', e.reason)
        html = None
        if num_retries > 0:
            if hasattr(e, 'code') and 500 <= e.code < 600:
                return download(url, user_agent, num_retries-1)
    return html
# 从excel中读取债券编码
def excel_read(filename):
    code_list = []
    data = xlrd.open_workbook(filename)
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    for rownum in range (0,nrows):
        rowdata = table.row_values(rownum)
        code_list.append(rowdata[0])
    return code_list


def get_date_rate(code_list):
    print code_list
    result = {}
    for code in code_list:
        url= 'http://www.chinabond.com.cn/jsp/include/EJB/fdllbd.jsp?sZqdm=%s'%code
        print code
        html = download(url)
        try:
            goal_list = re.findall(" <tr>\s+<td\s+width=9%\snowrap>[1-9][0-9]{0,2}</td>\s+<td width=15%\snowrap>(.*?)</td>\s+<td\swidth=15%\snowrap>(.*?)</td>\s+<td\swidth=12%\snowrap>(.*?)</td>\s+<td\swidth=12%\snowrap>(.*?)</td>\s+<td\s+width=12%\snowrap>(.*?)</td>",html)
        except:
            print url
        #goal_list = [goal for goal in goal_list if not goal==['&nbsp','&nbsp','&nbsp','&nbsp']
        #goal_list = re.findall('<td\s+class="dreport-row2">(.*?)</td>\s+<td\sclass="dreport-row2">(.*?)</td>\s+<td\s+class="dreport-row2">(.*?)</td>\s+<td\s+class="dreport-row2">(.*?)</td>\s+<td\sclass="dreport-row2">(.*?)</td>', html)
        result[code] = goal_list
    return result

#输出结果，写入excel中并保存
def excel_write(items):
    newTable = 'float_interest.xls'
    wb = xlwt.Workbook(encoding = 'utf8')
    ws = wb.add_sheet('result')
    headData = ['证券代码','本期起息日','本期结息日','基准利率','利差','本期利率']
    s = len(headData)
    for colnum in range(0,s):
        ws.write(0,colnum,headData[colnum],xlwt.easyxf('font:bold on'))
    index = 1
    for key in items:
        for item in items[key]:
            lens = len(item)
            ws.write(index,0,key)
            for i in range(1,lens+1):
                ws.write(index,i,item[i-1])
            index +=1
    wb.save(newTable)

filename ="ABS.xls"
code_list = excel_read(filename)
items = get_date_rate(code_list)
excel_write(items)

