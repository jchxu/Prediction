# coding=utf-8
#from WindPy import w
from functions import *

### 初始设置 ###
filename = "RB01_2017.xlsx"
TotalMoney = 1000000    #初始资金100万
TotalPos = 0            #初始仓位0
LastPX = 2900           #初始仓位最后成交价 PX=Price
LastACT = 0            #上一次开平标识符，1为开（买入），-1为平（卖出）
rate = 0.0007        #交易手续费率，万分之七
#timeinterval = 1     #每隔1分钟运行一次，单位为秒
MAFlag = [0,0,0]
MAjudge = 0
#####################

DayFileName = filename.split('.')[0]+"_Day_Records.txt"
DayFile=open(file=DayFileName, mode='w')
print(u'日期',u'资金',u'持仓',u'总资产',u'资金变化',u'持仓变化',u'持仓价值变化',u'总资产当日变化',u'总资产累积变化',sep=',',file=DayFile)
LogFileName = filename.split('.')[0]+"_All_Detail.txt"
LogFile = open(file=LogFileName, mode='w')
print(u'日期',u'交易价',u'5日平均',u'10日平均',u'20日平均',u'策略手数',u'实际手数',u'当前资金',u'当前持仓',sep=',',file=LogFile)

OriginAsset = TotalMoney + TotalPos * LastPX
HistoryData = ReadHistoryData(filename)     #['RB1701', '20170103', '0900', 2923.0]
HistoryDayPX = getDayPX(HistoryData)    #每日最后交易价为收盘价，[361, '20170103', 2920.0]，361为20170103最后一个分钟交易数据的index

for i in range(19, len(HistoryDayPX)):      #每天情况。前19天没有20日平均，所以从第20天开始
    NumAct = 0
    TodayMoney = TotalMoney
    TodayPos = TotalPos
    PreMA5, PreMA10, PreMA20 = calc_HistoryPreMA(HistoryDayPX,i)    #前19/9/4天收盘价总和
    MinuteDay = str(HistoryDayPX[i][1])[0:4]+'-'+str(HistoryDayPX[i][1])[4:6]+'-'+str(HistoryDayPX[i][1])[6:8]      #yyyy-mm-dd
    for j in range(HistoryDayPX[i-1][0]+1,HistoryDayPX[i][0]+1):            #有记录的i天所有分钟记录
        MinuteTime,MinutePX = getMinuteData(HistoryData[j])                     #获取分钟记录的时间和价格
        MA5, MA10, MA20 = calc_HistoryMA(PreMA5, PreMA10, PreMA20, MinutePX)    #根据分钟价格计算5/10/20日均价
        #MARecords.append([MinuteDay, MinuteTime, MinutePX, MA5, MA10, MA20])    #记录分钟价格及5/10/20日均价
        amount,MinuteMAFlag = calc_amount(MA5, MA10, MA20)                                   #根据日平均计算交易手数
        print("%s_%s,%.2f,%.2f,%.2f,%.2f,%4d" % (MinuteDay,MinuteTime,MinutePX,MA5,MA10,MA20,amount),sep=',', file=LogFile, end='')
        if MAjudge == 0:
            MAAct,MAFlag,MAjudge = judgeMAFlag(MAFlag,MinuteMAFlag,MAjudge)
        else:
            MAAct = 1
        if MAAct == 1:
            ActFlag, LastACT = judgeStatus(LastACT,amount)
        else:
            ActFlag = 0
        if (ActFlag != 0) :#and (NumAct <=4) :
            ActualAmount = checkAmount(ActFlag,amount,MinutePX,TodayMoney,TodayPos)
            TodayMoney, TodayPos = Trade(ActFlag,ActualAmount,rate,MinutePX,TodayMoney, TodayPos)
            NumAct += 1
            print("%s,%4d,%12.2f,%d" % ("",ActFlag*ActualAmount,TodayMoney,TodayPos),sep=',',file=LogFile)
            #print("%s_%s PX:%.2f MA5/10/20: %.2f/%.2f/%.2f Action: %4d MoneyNow:%12.2f PosNow:%d" % (MinuteDay,MinuteTime,MinutePX,MA5,MA10,MA20,ActFlag*ActualAmount,TodayMoney,TodayPos))
        elif (ActFlag == 0) :#or (NumAct >4):
            print("%s,%4s,%12.2f,%d" % ("","----",TodayMoney,TodayPos),sep=',',file=LogFile)
            #print("%s_%s PX:%.2f MA5/10/20: %.2f/%.2f/%.2f Action: %4s MoneyNow:%12.2f PosNow:%d" % (MinuteDay,MinuteTime,MinutePX,MA5,MA10,MA20,"----",TodayMoney,TodayPos))
    DayProfit = calc_profit(TotalMoney, TotalPos, LastPX, TodayMoney, TodayPos, HistoryDayPX[i][2],OriginAsset)
    PrintDayProfit(MinuteDay,DayProfit,DayFile)

    TotalMoney = TodayMoney
    TotalPos = TodayPos
    LastPX = HistoryDayPX[i][2]

LogFile.close()
DayFile.close()

#PXMAplot(LogFileName)
#DayProfitplot(DayFileName)
Plot_MA_Profit(LogFileName,DayFileName)
