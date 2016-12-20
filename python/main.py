# -*- coding: utf-8 -*-
#解决中文乱码问题
import sys
reload(sys)
sys.setdefaultencoding( "utf-8")

import xlrd, xlwt
import datetime
import json

masterTable = 'basicData.xls'
slaveTable= 'superShortData.xlsx'
newTable = 'iwtly.xls' 
slaveTable_mark= 'superShortData_mark.xls'

#将主从表交集部分写入新表
def excel_intersection (masterTable,slaveTable,newTable,begin_index = 0, by_index=0):
    slaveCodeList = getCodeList(slaveTable, by_index)  #获取超短期表的股票代码列表
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
        stockCode = rowData[0].split('.')[0]  #如果有. , 切出.之前的证劵代码
        if stockCode in slaveCodeList:#写入基础资料表中在 超短期表里 存在的一行
            for colnum in range(0, ncols):#遍历写入基础资料的那一行
                if table.cell(rownum, colnum).ctype == 3:  #如果是日期类型，进行datetime转换
                    rowData[colnum] = xlrd.xldate.xldate_as_datetime(rowData[colnum],0).strftime( '%Y-%m-%d' )
                ws.write(index, colnum, rowData[colnum])
            index = index + 1
    wb.save(newTable)#保存新表
#标记从表未匹配的行，方案是复制从表，设背景颜色
def  excel_intermark(slaveTable,newTable,begin_index = 0, by_index=0):
    newCodeList = getCodeList(newTable, by_index)#获取新表的股票代码列表
    slaveData = xlrd.open_workbook(slaveTable)#打开超短期表
    table = slaveData.sheets()[by_index]#取出超短期表的第一个标签内容
    nrows = table.nrows#行数
    ncols = table.ncols#列数
    wb = xlwt.Workbook();#新建新表对象
    ws = wb.add_sheet('iwtly');#新建新表标签
    headData = table.row_values(begin_index)#取出标题
    #设置写入格式风格，为了标记不匹配的内容，这里是蓝底加粗
    styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour ocean_blue; font: bold on;');
    for colnum in range(0, ncols):#写入头标题，并加粗
        ws.write(0, colnum, headData[colnum], xlwt.easyxf('font: bold on'))
    for rownum in range(1,nrows):#遍历超短期的每一行
        rowData = table.row_values(rownum)#取出改行的内容，返回是一列表
        stockCode = rowData[0].split('.')[0]#取出股票代码，如果有. ，切割出代码
        for colnum in range(0, ncols):#遍历此行的每一列
            if table.cell(rownum, colnum).ctype == 3: #如果是日期类型，进行datetime转换
                rowData[colnum] = xlrd.xldate.xldate_as_datetime(rowData[colnum],0).strftime( '%Y-%m-%d' )
            if not stockCode in newCodeList and colnum == 0: #如果股票代码不在新表，选中此行第一列，执行标记写入
                ws.write(rownum, colnum, rowData[colnum], styleBlueBkg)
            else:#其他正常写入
                ws.write(rownum, colnum, rowData[colnum])
    wb.save(slaveTable_mark)#保存标记表
#得到从表的证劵代码列表
def getCodeList(excelTable, by_index):
    codeList = []
    data = xlrd.open_workbook(excelTable)#打开表
    table = data.sheets()[by_index]#取出标签页内容
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    for rownum in range(1,nrows):#第二行开始遍历
        rowData = table.row_values(rownum)
        codeList.append(rowData[0].split('.')[0])#将股票代码添加入列表
    return codeList
#获取检测时间字典
def  excel_check(slaveTable,newTable,begin_index = 0, by_index=0): 
    data = xlrd.open_workbook(slaveTable)
    table = data.sheets()[by_index+1]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    check_dict={}
    for rownum in range(0,nrows):
        row_date,row_space,row_flag = table.row_values(rownum)
        if not check_dict.has_key(row_date):
            check_dict[row_date]={}
        check_dict[row_date][row_space]=row_flag
    return check_dict
#修正交易时间
def  excel_judge(slaveTable,newTable,check_dict,begin_index = 0, by_index=0): 
    data = xlrd.open_workbook(slaveTable)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    wb = xlwt.Workbook(); #新建新表对象
    ws = wb.add_sheet('iwtly');#新建新表标签
    for rownum in range(0,nrows):
        row_date,row_space = table.row_values(rownum)
        if not row_space in check_dict[row_date]:
            row_space = u'上海证券交易所'
        if check_dict[row_date][row_space] == '是':
            ws.write(rownum, 0, row_date)
        else: 
            for i in range(1,20): #20足够找到最近的工作日
                if check_dict[row_date+i][row_space] == '是':
                    ws.write(rownum, 0, row_date+i)
                    break
    wb.save(newTable)
    #print xlrd.xldate.xldate_as_datetime(row_date,0).strftime( '%Y-%m-%d' )
     
#主函数入口
def main():
    check_dict=excel_check('data1.xlsx','data2.xls',0,0)
    excel_judge('data1.xlsx','data2.xls',check_dict,0,0)
    #excel_intersection(masterTable,slaveTable,newTable,0, 0)
    #excel_intermark(slaveTable,newTable,0, 0)
#执行函数入口
if __name__=="__main__":
    main()
