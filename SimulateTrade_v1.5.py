# coding=utf-8
from WindPy import w
from xlsxwriter import Workbook
from functions import *
import time,sys

### 用户自定义设置 ###
TaxRate = 0.0007   #交易手续费率，万分之七
MARate = 0.01   #MA计算的浮动比例
OriginAsset = 1000000   #初始资金
timelist = ["8:00","23:59"]   #每天交易的时间段。开始，结束……，成对列出
timeinterval = 5   #程序运行的时间间隔，单位为秒
PlatformCode = '0000'   #交易/模拟平台（经纪商）代码
Account = 'W109839501202'   #交易/模拟账户名
Passwd = '000000'   #交易/模拟账户名密码
Kind = 'RB1901.SHF' #交易的商品品种
monthlist = [4,9,12]    #每年哪几个月的第二个星期五需要强制平仓
#####################

### 初始化 ###
AccountInfo = [PlatformCode,Account,Passwd]  #模拟账户平台/账号/*
LogFileName = GetDetailFileName()   #根据月份获取log文件的文件名，格式为yyyymm_Everyday_Detail.txt
DayFileName = "Day_Close.txt"   #记录每天摘要结果的文件名
LastAct = 0 #上次操作的标识，1为开仓，-1为平仓
ActFlag = 0 #当天操作标识，1为开仓，-1为平仓，-2为强制全部平仓，-3为强制平仓后的休整期,2为休整后的开仓
timelist2 = TransTimeList(timelist) #转换时间节点列表为datetime格式的列表
w.start()  #打开Wind-Python接口
Money,Pos,LastPX,LastAct,OriginMAFlag = ReadDayData(DayFileName,OriginAsset,Kind)   #获取昨日收盘后信息：资金，持仓，收盘价，上次操作标识,均值交叉标识
LogFile = CheckLogFile(LogFileName) #获取Log文件
TodayMoney = Money  #记录当天初始资金
TodayPos = Pos  #记录当天初始仓位
NumAct = 0  #每天交易次数
PreMA5, PreMA10, PreMA20 = GetClosePX(DayFileName,Kind)  #前4/9/19天收盘价预先求和

### 开始当天操作,检查日期 ###
TodayDate = str(datetime.date(datetime.now()))  #获取当天时间，格式为yyyy-mm-dd
#TodayDate = "2018-09-15"
ActFlag = CheckFriday(TodayDate,monthlist)  #检查当天是否强制平仓(-2)，或在休整期(-3),或继续(0)
print ("%s Records of %s %s" % ('-'*10,TodayDate,'-'*10))   #分割线

##for i in range(5): #测试运行n次
while CheckTime(timelist2):     #检查当天是否已收盘
    interval = CheckInterval(timeinterval,timelist2)    #获取循环操作停顿的时间间隔
    LivePX = w.wsq(Kind, "rt_latest").Data[0][0]  #获取最新成交价
    MA5,MA10,MA20 = CalcMA(PreMA5,PreMA10,PreMA20,LivePX)    #根据实时价格和预先求和，计算MA5，MA10，MA20
    NowTime = datetime.now().strftime('%Y-%m-%d_%H:%M:%S')  #获取当前时间，格式为2018-07-27_14:00:00
    PrintLiveLog(NowTime,LivePX,MA5,MA10,MA20,LogFile)
    if ActFlag == -2:  #指定月份第二个星期五，强制平仓
        TodayMoney,TodayPos,LastPX,LastAct,OriginMAFlag = ForceSell(Kind,AccountInfo,LivePX,TodayMoney,TodayPos,TaxRate)
        PrintTradeLog(ActFlag,Pos,TodayMoney,TodayPos,LogFile)
        continue
    elif ActFlag == -3:  # 尚在休整，仅记录当天收盘价后，直接退出
        continue
    else:   #非强制平仓、非休整期
        Amount,MAFlag = CalcAmount(MA5,MA10,MA20,MARate)  #根据5/10/20日均值，计算开平仓数量、均值交叉标识
        ActFlag,ActualAmount = CheckFlag(LastAct,Amount,OriginMAFlag,MAFlag,NumAct,LivePX,TodayMoney,TodayPos)
        if ActFlag != 0:    #Act标识可以进行操作
            #TodayMoney, TodayPos = Trade(ActFlag, ActualAmount, TaxRate, LivePX, TodayMoney, TodayPos)
            TodayMoney, TodayPos = SimulateTrade(Kind,AccountInfo,ActFlag,ActualAmount,TaxRate,LivePX,TodayMoney,TodayPos)   #模拟交易
            NumAct += 1     #当天交易次数加1
            LastAct = ActFlag   #更新ACT标识
            OriginMAFlag = MAFlag #更新均线交叉标识
        PrintTradeLog(ActFlag, ActualAmount, TodayMoney, TodayPos, LogFile)
        if CheckProfit(TodayMoney, TodayPos,LivePX,OriginAsset): #检查利润是否达到100%
            TodayMoney,TodayPos,LastPX,LastAct,OriginMAFlag = ForceSell(Kind,AccountInfo,LivePX,TodayMoney,TodayPos,TaxRate)
    time.sleep(interval)    #根据计算得到的时间间隔，暂停程序运行
ClosePX = w.wsd(Kind, "close", "ED0TD", date.today(), "Fill=Previous").Data[0][0]  # 已过收盘时间，获取获取当天收盘价
profit = calc_profit(Money, Pos, LastPX, TodayMoney, TodayPos, ClosePX, OriginAsset)  # 计算资金/持仓变化
print ("%s Summary of %s %s" % ('-'*10,TodayDate,'-'*10))    #分割线
DayFile = CheckDayFile(DayFileName, TodayDate)  #
PrintDayRecord(TodayDate,TodayMoney, TodayPos, ClosePX, profit, LastAct, DayFile)
#Money = TodayMoney
##Pos = TodayPos
#
LogFile.close()
DayFile.close()
w.stop()    #关闭Wind-Python接口
#
#Plot_Live_MAProfit(LogFileName, DayFileName)    #根据记录的Day、Log文件绘图