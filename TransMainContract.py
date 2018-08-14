# coding=utf-8

import xlrd
import numpy as np
from datetime import datetime
import pandas as pd

#sheetlist = [201601,201602,201603,201604,201605,201606,201607,201608,201609,201610,201611,201612,201701,201702,201703,201704,201705,201706,201707,201708,201709,201710,201711,201712]

#data1 = pd.read_csv('201601.csv',encoding='utf-8')
#print(data1)

#file = xlrd.open_workbook("RB_2016_2017.xlsx")
splitlist = ['RB1601','2016-05-01','RB1605','2016-10-01','RB1610','2017-01-01','RB1701','2017-05-01','RB1705','2017-10-01','RB1710']

ContractID = []
DateTime = []
Price = []

#for i in range(0,file.nsheets):
#    sheet = file.sheet_by_index(i)
#    for j in range(1,sheet.nrows):
#        ContractID.append(sheet.cell_value(j,0))
#        OriginT = xlrd.xldate.xldate_as_datetime(sheet.cell_value(j,1),0)
#        DateTime.append(OriginT)
#        Price.append(sheet.cell_value(j,5))
#    file.unload_sheet(i)
#print(list(set(ContractID)))#,DateTime,Price)

list1 = ['RB1708', 'RB1602', 'RB1611', 'RB1710', 'RB1809', 'RB1810', 'RB1605', 'RB1801', 'RB1803', 'RB1612', 'RB1812', 'RB1709', 'RB1608', 'RB1601', 'RB1610', 'RB1704', 'RB1603', 'RB1604', 'RB1703', 'RB1808', 'RB1806', 'RB1712', 'RB1705', 'RB1706', 'RB1607', 'RB1606', 'RB1805', 'RB1802', 'RB1701', 'RB1702', 'RB1609', 'RB1804', 'RB1811', 'RB1807', 'RB1707', 'RB1711']

list1.sort()
print(list1)