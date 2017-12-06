# -*- coding: utf-8 -*-

import sys
import datetime
#import xlrd
#from xlwt import *
import pandas as pd


reload(sys)
sys.setdefaultencoding('utf-8')

Total_Class = 8
File_Name = 'OA-HRIS_CKSH_WK46_2017_11_3(复件)'

print "Start" + datetime.time().__str__()

#w = Workbook()
#Sheet = 1
#while (Sheet <= Total_Class):
#    ws = w.add_sheet(u'班级' + Sheet.__str__())
#    Sheet = Sheet + 1

#w.save(File_Name + '_Class.xls')

DF = pd.read_excel(File_Name + ".xlsx")

#筛选出需要的列：   工号，员工姓名，在职状态，品牌，部门名称
DF2 = DF.filter(items=[u'工号', u'员工姓名', u'在职状态', u'品牌', u'部门名称'])
#print(DF2)

#筛选在职状态为‘在职’的
DF3 = DF2[DF2[u'在职状态'].isin([u'在职'])]
#print(DF3)

#筛选部门名称为 上海 苏州 的, 并根据‘部门名称’排序
DF4 = DF3[DF3[u'部门名称'].str.startswith("SH-") | DF3[u'部门名称'].str.startswith("SUZ-")]
#print(df4.sort_values(u'部门名称'))

#根据排序进行分课，课程数量Total_Class
DF5 = DF4.reset_index().filter(items=[u'工号', u'员工姓名', u'在职状态', u'品牌', u'部门名称'])
DF5.insert(5, 'Class', DF5.index % Total_Class + 1)
print DF5

writer = pd.ExcelWriter(File_Name + '_Class.xls')
DF5.to_excel(writer,u'完整名单')
Sheet = 1
while (Sheet <= Total_Class):
    Sheet_name = str(Sheet)
    DF_Sub = DF5[DF5[u'Class'].isin([Sheet])]
    print DF_Sub
    DF_Sub.to_excel(writer,u'班级'+ Sheet_name)
    Sheet = Sheet + 1
writer.save()
