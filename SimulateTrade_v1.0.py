# coding=utf-8
from WindPy import w
from xlsxwriter import Workbook
from functions import *
import time

### 用户自定义设置 ###
TaxRate = 0.0007   #交易手续费率，万分之七
MARate = 0.01   #MA计算的浮动比例
OriginAsset = 1000000   #初始资金
timelist = ["11:17","15:20"]   #每天交易的时间段。开始，结束……，成对列出
timeinterval = 5   #程序运行的时间间隔，单位为秒
PlatformCode = '0000'   #交易/模拟平台（经纪商）代码
Account = 'W109839501202'   #交易/模拟账户名
Passwd = '000000'   #交易/模拟账户名密码
Kinds = ['RB1808','RB1901']  #交易的商品品种列表
Kind = 'RB1901.SHF' #交易的商品品种
monthlist = [4,9,12]
#####################
w.start()   #打开Wind-Python接口

AccountInfo = [PlatformCode,Account,Passwd]  #模拟账户平台/账号/*
LogFileName = getDetailFileName()   #根据月份获取log文件的文件名，格式为yyyymm_Everyday_Detail.txt
DayFileName = "Day_Close.txt"   #记录每天摘要结果的文件名
TodayDate = str(datetime.date(datetime.now()))  #获取当天时间，格式为yyyy-mm-dd
OriginMAFlag = [0,0,0]    ##均值（MA）交叉标识，用于判断是否有MA的交叉情况
MAjudge = 0 #判断
LastACT = 0 #上一次操作的标识符，1为开仓，-1为平仓

Money,Pos,LastPX,LastACT = ReadDayData(DayFileName,OriginAsset,Kind)   #获取昨日收盘后信息：资金，持仓，收盘价，上次操作标识
LogFile = CheckLogFile(LogFileName) #获取Log文件
TodayMoney = Money
TodayPos = Pos
NumAct = 0  #每天交易次数

w.start()   #打开Wind-Python接口
PreMA5, PreMA10, PreMA20 = getClosePX(DayFileName,Kind)  #前4/9/19天收盘价预先求和
print ("%s Records of %s %s" % ('-'*10,TodayDate,'-'*10))   #分割线
timelist2 = TransTimeList(timelist) #转换时间节点列表为datetime格式的列表
#for i in range(5): #测试运行n次
while CheckTime(timelist2):     #检查当天是否已收盘
    interval = CheckInterval(timeinterval,timelist2)    #获取循环操作停顿的时间间隔
    LivePX = w.wsq(Kind, "rt_latest").Data[0][0]  #获取最新成交价
    MA5, MA10, MA20 = calc_MA(PreMA5, PreMA10, PreMA20, LivePX)    #根据实时价格和预先求和，计算MA5，MA10，MA20
    NowTime = datetime.now().strftime('%Y-%m-%d_%H:%M:%S')  #获取当前时间，格式为2018-07-27_14:00:00
    PrintLiveLog(NowTime,LivePX,MA5,MA10,MA20,LogFile)
    amount,MinuteMAFlag = calc_amount(MA5, MA10, MA20,MARate)  #根据5/10/20日均值，计算开平仓数量、均值交叉标识
    ActFlag,ActualAmount = checkFlag(LastACT, amount, OriginMAFlag, MinuteMAFlag, NumAct, monthlist)
    if ActFlag != 0:    #Act标识可以进行操作
        ActualAmount = checkAmount(ActFlag, amount, LivePX, TodayMoney, TodayPos)   #计算真实应该交易的数量
        if ActualAmount != 0:
            # TodayMoney, TodayPos = Trade(ActFlag, ActualAmount, TaxRate, LivePX, TodayMoney, TodayPos)
            TodayMoney, TodayPos = SimulateTrade(Kind,AccountInfo, ActFlag, ActualAmount, TaxRate, LivePX, TodayMoney, TodayPos)   #模拟交易
            NumAct += 1     #当天交易次数加1
            LastACT = ActFlag   #更新ACT标识
            OriginMAFlag = MinuteMAFlag #更新均线交叉标识
    PrintTradeLog(ActFlag, ActualAmount, TodayMoney, TodayPos, LogFile)

    time.sleep(interval)    #根据计算得到的时间间隔，暂停程序运行
ClosePX = w.wsd(Kind, "close", "ED0TD", date.today(), "Fill=Previous").Data[0][0]  # 获取获取今日收盘价
profit = calc_profit(Money, Pos, LastPX, TodayMoney, TodayPos, ClosePX, OriginAsset)  # 计算资金/持仓变化
print ("%s Summary of %s %s" % ('-'*10,TodayDate,'-'*10))    #分割线
DayFile = CheckDayFile(DayFileName, TodayDate)  #
PrintDayRecord(TodayDate,TodayMoney, TodayPos, ClosePX, profit, LastACT, DayFile)
#Money = TodayMoney
#Pos = TodayPos

LogFile.close()
DayFile.close()
w.stop()    #关闭Wind-Python接口

Plot_Live_MAProfit(LogFileName, DayFileName)    #根据记录的Day、Log文件绘图