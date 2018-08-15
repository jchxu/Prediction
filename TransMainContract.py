# coding=utf-8

import xlrd
import numpy as np
import datetime
import pandas as pd

#start = datetime.datetime.now()
#file = xlrd.open_workbook("RB_2016_2017.xlsx")
splitlist = ['RB1605','2016-03-15','RB1610','2016-08-20','RB1701','2016-12-01','RB1705','2017-03-15','RB1710','2017-08-20','RB1801','2017-12-01','RB1805']

### 获取所有数据的分列表，品种/时间/最新价格
#ContractID = []
#DateTime = []
#Price = []
#for i in range(0,file.nsheets):
#    sheet = file.sheet_by_index(i)
#    for j in range(1,sheet.nrows):
#        ContractID.append(sheet.cell_value(j,0))
#        OriginT = xlrd.xldate.xldate_as_datetime(sheet.cell_value(j,1),0)
#        DateTime.append(OriginT)
#        Price.append(sheet.cell_value(j,5))
#    file.unload_sheet(i)
#print(ContractID)
#print(DateTime)
#print(Price)

### 每个品种单独保存为一个文件
#UniqueContract = list(set(ContractID))  # ['RB1601', 'RB1602', 'RB1603', 'RB1604', 'RB1605', 'RB1606', 'RB1607', 'RB1608', 'RB1609', 'RB1610', 'RB1611', 'RB1612', 'RB1701', 'RB1702', 'RB1703', 'RB1704', 'RB1705', 'RB1706', 'RB1707', 'RB1708', 'RB1709', 'RB1710', 'RB1711', 'RB1712', 'RB1801', 'RB1802', 'RB1803', 'RB1804', 'RB1805', 'RB1806', 'RB1807', 'RB1808', 'RB1809', 'RB1810', 'RB1811', 'RB1812']
#FileDict = {}
#for i in range(0,len(UniqueContract)):
#    file = open(UniqueContract[i]+".csv", mode='w')
#    print('ContractID','DateTime','LastPX',sep=',',file=file)
#    FileDict[UniqueContract[i]] = file
#for i in range(0,len(ContractID)):
#    print(ContractID[i],DateTime[i],Price[i],sep=',',file=FileDict[ContractID[i]])
#for i in range(0,len(UniqueContract)):
#    FileDict[UniqueContract[i]].close()
#end = datetime.datetime.now()
#print((end-start).seconds)

### 根据列表拆分并组合文件
file = open('MainContract.csv', mode='w')
print('ContractID','DateTime','LastPX',sep=',',file=file)
file.close()
for i in range(0,len(splitlist),2):
    if i == 0:
        csvfile = pd.read_csv(splitlist[i]+'.csv')
        data = csvfile.loc[csvfile['DateTime'] < (splitlist[i+1]+" 00:00:00")]
        data.to_csv("MainContract.csv", header=False, index=False, mode='a')
    elif i < (len(splitlist)-1):
        csvfile = pd.read_csv(splitlist[i] + '.csv')
        data = csvfile.loc[((splitlist[i-1] + " 00:00:00") <= csvfile['DateTime']) & (csvfile['DateTime'] < (splitlist[i+1] + " 00:00:00"))]
        data.to_csv("MainContract.csv", header=False, index=False, mode='a')
    elif i == (len(splitlist)-1):
        csvfile = pd.read_csv(splitlist[i] + '.csv')
        data = csvfile.loc[(splitlist[i-1] + " 00:00:00") <= csvfile['DateTime']]
        data.to_csv("MainContract.csv", header=False, index=False, mode='a')
