# coding=utf-8
#from WindPy import w
from xlsxwriter import Workbook
from functions import *
import time
import random
import matplotlib.pyplot as plt

### 初始设置 ###
filename = "RB01_2017Q1.xlsx"
TotalMoney = 1000000    #初始资金100万
TotalPos = 0            #初始仓位0
LastPX = 2900           #初始仓位最后成交价
LastACT = -1            #开平标识符，1为开，-1为买
rate = 0.0007        #交易手续费率，万分之七
timeinterval = 1     #每隔1分钟运行一次，单位为秒
Records = []         #交易记录
MARecords = []       #日均值记录
#####################

OriginAsset = TotalMoney + TotalPos * LastPX
HistoryData = ReadHistoryData(filename)
#print(HistoryData)
HistoryDayPX = getDayPX(HistoryData)    #以每日最后交易价为收盘价
#print(HistoryDayPX)
HistoryDayStatus = filename.split('.')[0]+"_HistoryDayStatus.txt"
f1=open(HistoryDayStatus,mode='w')
logfilename = filename.split('.')[0]+"_Detail.log"
logfile = open(file=logfilename, mode='w')
print(u'历史日期',u'当天资金',u'资金变化',u'当日持仓',u'持仓变化',u'持仓价值',u'持仓价值变化',u'总资产',u'总资产当日变化',u'总资产净值',sep=',',file=f1)

for i in range(19, len(HistoryDayPX)):
    TodayMoney = TotalMoney
    TodayPos = TotalPos
    PreMA5, PreMA10, PreMA20 = calc_HistoryPreMA(HistoryDayPX,i)
    #print(HistoryDayPX[i][1],PreMA5, PreMA10, PreMA20)
    for j in range(HistoryDayPX[i-1][0]+1,HistoryDayPX[i][0]+1):
        HistoryMin1Data = HistoryData[j]
        HistoryMin1PX = HistoryMin1Data[3]
        HistoryMin1Day = HistoryMin1Data[1]
        HistoryMin1Time = HistoryMin1Data[2]
        Min1DayTime = HistoryMin1Day+' '+HistoryMin1Time
        MA5, MA10, MA20 = calc_HistoryMA(PreMA5, PreMA10, PreMA20, HistoryMin1PX)
        MARecords.append([HistoryMin1Day, HistoryMin1Time, HistoryMin1PX, MA5, MA10, MA20])
        print(HistoryMin1Day, HistoryMin1Time, HistoryMin1PX, MA5, MA10, MA20,sep=',', file= logfile, end='')
        amount = judge_status(MA5, MA10, MA20)
        TodayMoney, TodayPos, Records = HistoryTrade(LastACT, rate, HistoryMin1PX, Min1DayTime, amount, TodayMoney, TodayPos, Records)
    profit = calc_profit(TotalMoney, TotalPos, LastPX, TodayMoney, TodayPos, HistoryDayPX[i][2])
    print(HistoryDayPX[i][1],TodayMoney,profit[0],TodayPos,profit[1],TodayPos*HistoryDayPX[i][2],profit[2],TodayMoney+TodayPos*HistoryDayPX[i][2],profit[3],TodayMoney+TodayPos*HistoryDayPX[i][2]-OriginAsset,sep=',',file=f1)
    #printprofit(profit)
    TotalMoney = TodayMoney
    TotalPos = TodayPos
    LastPX = HistoryDayPX[i][2]

Historyfilename = filename.split('.')[0]+"_HistoryPX&MA.txt"
f2=open(Historyfilename,mode='w')
print(u'历史日期',u'当天时间',u'交易价格',u'5日平均',u'10日平均',u'20日平均',sep=',',file=f2)
for i in range(0,len(MARecords)):
    print(MARecords[i][0],MARecords[i][1],MARecords[i][2],MARecords[i][3],MARecords[i][4],MARecords[i][5],sep=',',file=f2)
logfile.close()
f1.close()
f2.close()
MAplot(MARecords)
