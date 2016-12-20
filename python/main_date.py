# -*- coding: utf-8 -*-
#解决中文乱码问题
import sys
reload(sys)
sys.setdefaultencoding( "utf-8")

import xlrd, xlwt
import datetime, time
import json

masterTable = 'abs-money.xls'
slaveTable= 'abs-money.xls'
newTable = 'abs-money-output.xls'

def excel_intersection (masterTable,slaveTable,newTable,begin_index = 0, by_index=0):
    slave_list = getslaveData(slaveTable, by_index+1)  #获取从表的资源数据（第二标签页）
    masterData = xlrd.open_workbook(masterTable) #打开基础资料表
    wb = xlwt.Workbook(); #新建新表对象
    ws = wb.add_sheet('iwtly');#新建新表标签
    table = masterData.sheets()[by_index]#取出主表的第一个标签页
    nrows = table.nrows#行数
    ncols = table.ncols#列数
    headData = table.row_values(begin_index)#excel的标题获取
    index = 1  #写入的索引，依次写入每一行
    for colnum in range(0, ncols):#在新表中写入头标题，并加粗
        ws.write(0, colnum, headData[colnum], xlwt.easyxf('font: bold on'))
    for rownum in range(1,nrows):#第二行开始遍历
        rowData = table.row_values(rownum)
        count = checktime(rowData[7], rowData[8], slave_list)
        rowData[10] = count if count > 0 else ''
        while(count >= 0):
            for colnum in range(0, ncols):
                if table.cell(rownum, colnum).ctype == 3:  #如果是日期类型，进行datetime转换(只转换一次)
                    temp = xlrd.xldate.xldate_as_datetime(rowData[colnum],0).strftime( '%Y-%m-%d' )
                    ws.write(index, colnum, temp)
                else:
                    ws.write(index, colnum, rowData[colnum])
            count -= 1
            index += 1
    wb.save(newTable)#保存新表

def getslaveData(excelTable, by_index):
    data = xlrd.open_workbook(excelTable)#打开表
    table = data.sheets()[by_index]#取出标签页内容
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    slave_list = []
    for rownum in range(1,nrows):#第二行开始遍历
        rowData = table.row_values(rownum)
        slave_list.append(int(time.mktime(
                        time.strptime("%s 00:00:00" %rowData[0], "%Y-%m-%d  %H:%M:%S"))))
    return slave_list

def checktime(starttime,endtime,slave_list):
    count = 0
    try:
        starttime = xlrd.xldate.xldate_as_datetime(starttime,0).strftime( '%Y-%m-%d' )
        starttime = int(time.mktime(time.strptime("%s 00:00:00" %starttime, "%Y-%m-%d  %H:%M:%S")))
    except:
        starttime = int(time.mktime(time.strptime("%s 00:00:00" %starttime, "%Y/%m/%d  %H:%M:%S")))
        pass
    try:
        endtime = xlrd.xldate.xldate_as_datetime(endtime,0).strftime( '%Y-%m-%d' )
        endtime = int(time.mktime(time.strptime("%s 00:00:00" %endtime, "%Y-%m-%d  %H:%M:%S")))
    except:
        endtime = int(time.mktime(time.strptime("%s 00:00:00" %endtime, "%Y/%m/%d  %H:%M:%S")))
        pass

    #endtime = xlrd.xldate.xldate_as_datetime(endtime,0).strftime( '%Y-%m-%d' )
    for stime in slave_list:
        if starttime <= stime and endtime >= stime:
            count += 1
    return count
#主函数入口
def main():
    excel_intersection(masterTable,slaveTable,newTable,0, 0)
    #excel_intermark(slaveTable,newTable,0, 0)
#执行函数入口
if __name__=="__main__":
    main()
