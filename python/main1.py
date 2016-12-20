# -*- coding: utf-8 -*-
#解决中文乱码问题
import sys
reload(sys)
sys.setdefaultencoding( "utf-8")

import xlrd, xlwt
import datetime
import json

masterTable = 'dataSource.xls'
slaveTable= 'dataSource.xls'
newTable = 'dataOutput.xls'

def excel_intersection (masterTable,slaveTable,newTable,begin_index = 0, by_index=0):
    slave_dict = getslaveData(slaveTable, by_index+1)  #获取从表的资源数据（第二标签页）
    masterData = xlrd.open_workbook(masterTable) #打开基础资料表
    wb = xlwt.Workbook(); #新建新表对象
    ws = wb.add_sheet('iwtly');#新建新表标签
    styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour ocean_blue; font: bold on;');  #蓝色样式
    styleRedBkg = xlwt.easyxf('pattern: pattern solid, fore_colour red; font: bold on;');#红色标示
    styleGreenBkg = xlwt.easyxf('pattern: pattern solid, fore_colour green; font: bold on;');#绿色标示
    table = masterData.sheets()[by_index]#取出主表的第一个标签页
    nrows = table.nrows#行数
    ncols = table.ncols#列数
    headData = table.row_values(begin_index)#excel的标题获取
    index = 1  #写入的索引，依次写入每一行
    for colnum in range(0, ncols):#在新表中写入头标题，并加粗
        ws.write(0, colnum, headData[colnum], xlwt.easyxf('font: bold on'))
    for rownum in range(1,nrows):#第二行开始遍历
        rowData = table.row_values(rownum)
        stockCode = rowData[0]  #如果有. , 切出.之前的证劵代码
        if slave_dict.has_key(stockCode) :#写入主表中在 从表里 存在的一行
            colnum = 0
            for key in slave_dict[stockCode]:
                tmp = slave_dict[stockCode][key]
                if table.cell(rownum, colnum).ctype == 3:
                    rowData[colnum] = xlrd.xldate.xldate_as_datetime(rowData[colnum],0).strftime( '%Y-%m-%d' )
                if not rowData[colnum] == '' and tmp == '':
                    ws.write(index, colnum, rowData[colnum], styleGreenBkg)
                elif not rowData[colnum] == '' and not rowData[colnum] == tmp:
                    ws.write(index, colnum, tmp, styleBlueBkg)
                else:
                    ws.write(index, colnum, rowData[colnum])
                colnum += 1
        else:  #如果主表的债券代码在从表不存在，标记
            for colnum in range(0, ncols):#遍历写入主表的那一行
                if table.cell(rownum, colnum).ctype == 3:  #如果是日期类型，进行datetime转换
                    rowData[colnum] = xlrd.xldate.xldate_as_datetime(rowData[colnum],0).strftime( '%Y-%m-%d' )
                if colnum == 0:
                    ws.write(index, colnum, rowData[colnum], styleRedBkg)
        index = index + 1
    wb.save(newTable)#保存新表

def getslaveData(excelTable, by_index):
    data = xlrd.open_workbook(excelTable)#打开表
    table = data.sheets()[by_index]#取出标签页内容
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    slave_dict = {}
    for rownum in range(1,nrows):#第二行开始遍历
        rowData = table.row_values(rownum)
        for colnum in range(1, ncols):
            if table.cell(rownum, colnum).ctype == 3:  #如果是日期类型，进行datetime转换
                rowData[colnum] = xlrd.xldate.xldate_as_datetime(rowData[colnum],0).strftime( '%Y-%m-%d' )
        slave_dict[rowData[0]] = {'f0': rowData[0],'f1': rowData[1],'f2': rowData[2],'f3': rowData[3],
                                  'f4': rowData[4],'f5': rowData[5],'f6': rowData[6],'f7': rowData[7],
                                  'f8': rowData[8],'f9': rowData[9],'f10': rowData[10],'f11': rowData[11]}
    return slave_dict

#主函数入口
def main():
    excel_intersection(masterTable,slaveTable,newTable,0, 0)
    #excel_intermark(slaveTable,newTable,0, 0)
#执行函数入口
if __name__=="__main__":
    main()
