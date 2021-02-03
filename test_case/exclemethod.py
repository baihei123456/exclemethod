# -*- coding: UTF-8 -*-
import xlrd
from xlutils import copy
def get_excleDate(sheetName,startRow,endRow,body=4,repsData=5):
    resList=[]
    excleDir='../date_excle/1登录测试用例.xls'
    '''打开excle'''
    workBook=xlrd.open_workbook(excleDir)
    # sheets=workBook.sheet_names() #读取sheet
    # print(sheets)
    #取对应sheet里面的内容
    workSheet=workBook.sheet_by_name(sheetName)
    #return workSheet.row_values(1)#读取一行数据
    #读取单元格
    #print(workSheet.cell(2,6).value)
    #预期结果
    #print(workSheet.cell(1,6).value)
    for one in range(startRow-1,endRow):
        resList.append((workSheet.cell(one,body).value,workSheet.cell(one,repsData).value))
    return resList

get_excleDate('登录测试用例',2,12)
for t in get_excleDate('登录测试用例',2,12):
    print(t)

#写入excle
def write_excleDate():
    excleDir='../date_excle/1登录测试用例.xls'
    '''打开excle'''
    workBook=xlrd.open_workbook(excleDir,formatting_info=True)
    newWorkBook=copy(workBook)