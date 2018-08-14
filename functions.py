# coding=utf-8
from WindPy import w
from datetime import datetime, date, timedelta
import xlrd,time,os, operator, calendar
import matplotlib.pyplot as plt
import numpy as np

### 根据当前月份定义文件名
def GetDetailFileName():
    ThisMonth = str(datetime.now().strftime('%Y%m'))
    FileName = ThisMonth+"_Everyday_Detail.txt"
    return FileName

### 获取本年度所有指定月份的第二个星期五，返回列表
def GetFridays(monthlist):
    Fridays = []
    ThisYear = int(datetime.date(datetime.now()).year)
    for i in range(0, len(monthlist)):
        num = 0
        for j in range(1, calendar.monthrange(ThisYear, monthlist[i])[1] + 1):
            if datetime(year=ThisYear, month=monthlist[i], day=j).weekday() == calendar.FRIDAY:
                if num == 1:
                    Fridays.append(datetime(year=ThisYear, month=monthlist[i], day=j).strftime('%Y-%m-%d'))
                    break
                else:
                    num += 1
    return Fridays

### 检查是否在强制平仓后的休整期(强制平仓后的30天)
def CheckRest(TodayDate,Fridays):
    flag = 0
    TodayList = TodayDate.split('-')
    Today = datetime(year=int(TodayList[0]),month=int(TodayList[1]),day=int(TodayList[2]))
    for i in range(0,len(Fridays)):
        FridayList = Fridays[i].split('-')
        Friday = datetime(year=int(FridayList[0]), month=int(FridayList[1]), day=int(FridayList[2]))
        if (Today > Friday) and (Today <= (Friday+timedelta(days=30))):
            flag = 1
            break
        else:
            flag = 0
    if flag == 1: return True
    elif flag == 0: return False

### 检查当天是否强制平仓，或在休整期
def CheckFriday(TodayDate,monthlist):
    ActFlag = 0
    Fridays = GetFridays(monthlist)
    if TodayDate in Fridays:
        print(u"今日强制平仓！")
        ActFlag = -2
    elif CheckRest(TodayDate, Fridays):
        print(u"尚在强制平仓后的休整期！")
        ActFlag = -3
    return ActFlag

### 返回两个时间点之间的间隔秒数，至少为1秒
def ReturnInterval(time1, time2) :
    interval = 1
    if time1 == 0:
        nowtime = datetime.now()
        interval = max((time2-nowtime).seconds, 1)
    else:
        interval = max((time2-time1).seconds, 1)
    return interval

### 检查Day文件是否存在
def CheckDayFile(DayFileName, TodayDate):
    if os.path.exists(DayFileName):
        DayFile = open(DayFileName, mode='r')
        Lines = DayFile.readlines()
        NumL = len(Lines)
        DayFile.close()
        if NumL == 0:   #有Day文件，无内容，输出标题行，继续追加
            DayFile = open(DayFileName, mode='w')
            print(u"日期", u"资金", u"持仓", u"收盘", u"总资产", u"资金变化", u"持仓变化", u"收益日变化", u"累积收益", u"上次操作", u"上次均线",sep=',', file=DayFile)
        elif NumL == 1: #有Day文件，只有标题行，继续追加
            DayFile = open(DayFileName, mode='a')
        else:   #有Day文件，有内容，删除今天内容，继续追加
            templist = []
            for i in range(1, len(Lines)):
                linedate = Lines[i].split(',')[0]
                if linedate != TodayDate:
                    templist.append(Lines[i])
            DayFile = open(DayFileName, mode='w')
            print(u"日期", u"资金", u"持仓", u"收盘", u"总资产", u"资金变化", u"持仓变化", u"收益日变化", u"累积收益", u"上次操作",  u"上次均线",sep=',', file=DayFile)
            for j in range(0,len(templist)):
                print(templist[j].strip(), sep=',', file=DayFile)
    else:   #无Day文件，新建，输出标题行，继续追加
        DayFile = open(DayFileName, mode='w')
        print(u"日期", u"资金", u"持仓", u"收盘", u"总资产", u"资金变化", u"持仓变化", u"收益日变化", u"累积收益", u"上次操作",  u"上次均线",sep=',', file=DayFile)
    return DayFile

### 打印day文件，同时输出到屏幕
def PrintDayRecord(TodayDate,TodayMoney, TodayPos, ClosePX, profit, LastAct, DayFile):
    print(TodayDate,TodayMoney, TodayPos, ClosePX, TodayMoney+TodayPos*ClosePX, profit[3], profit[4], profit[6], profit[7], LastAct, [0,0,0], sep=',', file=DayFile)
    print(u"日期:%s, 资金:%r, 持仓:%r, 收盘:%r, 总资产:%r, 资金变化:%r, 持仓变化:%r, 收益日变化:%r, 累积收益:%r, 上次操作:%s, 上次均线:%s" % (TodayDate,TodayMoney, TodayPos, ClosePX, TodayMoney+TodayPos*ClosePX, profit[3], profit[4], profit[6], profit[7], LastAct, [0,0,0]))

### 根据Day文件或从Wind平台读取以往资金、持仓、收盘价、上次操作标识符
def ReadDayData(DayFileName,OriginAsset,Kind):
    flag = 0
    if os.path.exists(DayFileName):
        DayFile = open(DayFileName, mode='r')
        Lines = DayFile.readlines()
        DayFile.close()
        if len(Lines)>1:    #有DayFile文件且存在以往数据，从此文件最后一行读取以往数据
            lastline = Lines[-1].split(',')
            Money = float(lastline[1])
            Pos = int(lastline[2])
            LastPX = float(lastline[3])
            LastAct = int(lastline[-2])
            OriginMAFlag = lastline[-1]
        else:     #有DayFile文件但无以往数据（只有标题行），从Wind读取以往数据
            flag = 1
    else:   #无DayFile文件，从Wind读取以往数据
        flag = 1
    if flag == 1:
        Money = OriginAsset
        Pos = 0
        LastPXData = w.wsd(Kind, "close", "ED-1TD", datetime.today(), "Fill=Previous")
        if LastPXData.ErrorCode == -103:
            print(u"Wind客户端尚未打开并登录！")
            exit()
        else:
            LastPX = LastPXData.Data[0][0]
        LastAct = 0
        OriginMAFlag = [0,0,0]
    return Money,Pos,LastPX,LastAct,OriginMAFlag

### 检查Log文件是否存在
def CheckLogFile(LogFileName):
    if os.path.exists(LogFileName):
        LogFile = open(LogFileName, mode='r')
        NumL = len(LogFile.readlines())
        LogFile.close()
        if NumL > 0:    #有Day文件，有内容，准备追加输入
            LogFile = open(LogFileName, mode='a')
        else:   #有Day文件，但无内容，输入标题行，准备追加输入
            LogFile = open(LogFileName, mode='a')
            print(u"日期_时间", u"实时价格", u"5日平均",u"10日平均",u"20日平均", u"操作标识符", u"交易数量", u"资金", u"持仓", sep=',', file=LogFile)
    else:   #无Day文件，新建文件，输入标题行，准备追加输入
        LogFile = open(LogFileName, mode='w')
        print(u"日期_时间", u"实时价格", u"5日平均",u"10日平均",u"20日平均", u"操作标识符", u"交易数量", u"资金", u"持仓", sep=',', file=LogFile)
    return LogFile

### 根据Day文件或从Wind平台读取数据，前4/9/19天收盘价预先求和
def GetClosePX(DayFileName,Kind):
    flag = 0
    if os.path.exists(DayFileName):     #有Day文件
        DayFile = open(DayFileName, mode='r')
        Lines = DayFile.readlines()
        NumL = len(Lines)
        DayFile.close()
        if NumL >= 20:  #Day文件至少包含19天的数据，需要读取最后19行里的数据
            OldClosePX = np.loadtxt(fname=DayFileName, delimiter=',', skiprows=1, usecols=3, unpack=True)
            PreMA20 = sum(OldClosePX[NumL-19:NumL]) #前19天收盘价求和
            PreMA10 = sum(OldClosePX[NumL-9:NumL])  #前9天收盘价求和
            PreMA5 = sum(OldClosePX[NumL-4:NumL])   #前4天收盘价求和
        else:   #有Day文件，但数据量不够,需从Wind获取数据
            flag = 1
    else:   #无Day文件,需从Wind获取数据
        flag = 1
    if flag == 1:   #从Wind获取数据
        olddata = w.wsd(Kind, "close", "ED-19TD", datetime.today(), "Fill=Previous")    #从Wind读取前19天收盘价数据，返回结果为前19天+今天的价格
        Pre19Close = olddata.Data[0][0:19]   #前19天收盘价
        Pre9Close = olddata.Data[0][10:19]   #前9天收盘价
        Pre4Close = olddata.Data[0][15:19]   #前4天收盘价
        PreMA20 = sum(Pre19Close)
        PreMA10 = sum(Pre9Close)
        PreMA5 = sum(Pre4Close)
    return PreMA5, PreMA10, PreMA20

### 检查指定的时间节点是否为偶数个、按时间先后顺序，并转换为datetime格式
def TransTimeList(timelist):
    nowtime = datetime.now()
    timelist2 = []
    if (len(timelist)%2) != 0:  #时间节点应为偶数个。起始，结束，起始，结束，……
        print(u"时间范围首尾不配对！")
        exit()
    for i in range(0, len(timelist)):   #替换当前时间中的小时/分钟为指定的时间节点
        tt = timelist[i].split(':')
        timelist2.append(nowtime.replace(hour=int(tt[0]), minute=int(tt[1]), second=0))
    for i in range(0,len(timelist2)-1): #时间节点应该先后顺序排列
        if not (timelist2[i] < timelist2[i+1]):
            print(u"时间先后顺序不对",timelist2[i],timelist2[i+1])
            exit()
    return timelist2

### 检查当前时间是否已过当天收盘时间
def CheckTime(timelist2):
    nowtime = datetime.now()
    if nowtime > timelist2[-1]:     #默认时间列表中最后一个应为收盘时间点
        print(u"已过收盘时间：",timelist2[-1])
        return False
    else:
        return True

### 计算时间间隔
def CheckInterval(initinterval,timelist2):
    nowtime = datetime.now()
    if nowtime < timelist2[0]:  #当前时间还未到最早开盘时间，时间间隔为两者时间差秒数与1秒的最大值。（避免差0.*秒到开盘时间，时间差秒数为0的情况）
        print(u"尚未到开盘时间，等待开盘：",timelist2[0])
        interval = max((timelist2[0]-nowtime).seconds,1)
    else:
        for i in range(0, int(len(timelist2)/2)):   #以（开始-结束）时间对为索引
            if (timelist2[2*i] <= nowtime) and (nowtime <= timelist2[2*i+1]):   #在交易时间段内时，时间间隔为初始时间间隔
                interval = initinterval
                break
            elif (timelist2[2*i+1] < nowtime) and (nowtime < timelist2[2*i+2]): #在两个交易时间段之间时，时间间隔为当前时间与后一个开盘时间的时间差秒数与1秒的最大值
                interval = max((timelist2[2*i+2]-nowtime).seconds,1)
                break
    return interval



### 根据实时价格和预先求和，计算MA5，MA10，MA20
def CalcMA(PreMA5, PreMA10, PreMA20, LivePX):
    MA5 = (PreMA5 + LivePX) / 5
    MA10 = (PreMA10 + LivePX) / 10
    MA20 = (PreMA20 + LivePX) / 20
    #print(u'实时成交价：%s，5日均值：%s，10日均值：%s，20日均值：%s' % (LivePX, MA5, MA10, MA20))
    return (MA5, MA10, MA20)

### 打印log文件，同时输出到屏幕
def PrintLiveLog(NowTime,LivePX,MA5,MA10,MA20,LogFile):
    print(u"日期_时间:%s, 实时价格:%r, 5/10/20日平均:%r/%r/%r" % (NowTime, LivePX, MA5, MA10, MA20), end=',')
    print(NowTime, LivePX, MA5, MA10, MA20, sep=',', end=',', file=LogFile)

### 强制平仓
def ForceSell(Kind, AccountInfo,LivePX,TodayMoney, TodayPos,TaxRate):
    data1 = w.tlogon(BrokerID=AccountInfo[0], DepartmentID=0, LogonAccount=AccountInfo[1], Password=AccountInfo[2], AccountType='shf')  # 登录模拟账户
    data21 = w.torder(SecurityCode=Kind, TradeSide='Sell', OrderPrice=LivePX, OrderVolume=TodayPos, MarketType='SHF')  # 模拟平仓
    TodayMoney += LivePX*TodayPos*(1-TaxRate)
    TodayPos = 0
    LastAct = -2
    OriginMAFlag = [0,0,0]
    return TodayMoney, TodayPos, LivePX, LastAct, OriginMAFlag

### 打印log文件，同时输出到屏幕
def PrintTradeLog(ActFlag, ActualAmount, TodayMoney, TodayPos, LogFile):
    print(" 操作标识符:%s, 交易数量:%d, 资金:%r, 持仓:%r" % (ActFlag, ActualAmount, TodayMoney, TodayPos))
    print(ActFlag, ActualAmount, TodayMoney, TodayPos, sep=',', file=LogFile)

### 根据5/10/20日均值，计算开平仓数量、均值交叉标识
def CalcAmount(MA5, MA10, MA20,MARate):    #三个判断条件各自独立，最终开平仓结果为三个判断条件的结果求和
    Amount = 0      #开/平仓的具体手数
    MAFlag = [0,0,0]  #MA均值交叉标识
    #条件1
    if MA10 > MA20*(1+MARate):
        Amount += 120
        MAFlag[0] = 1
    elif MA10 < MA20*(1-MARate):
        Amount += -150
        MAFlag[0] = -1
    # 条件2
    if MA5 > MA20*(1+MARate):
        Amount += 80
        MAFlag[1] = 1
    elif MA5 < MA20*(1-MARate):
        Amount += -80
        MAFlag[1] = -1
    # 条件3
    if MA5 > MA10*(1+MARate):
        Amount += 50
        MAFlag[2] = 1
    elif MA5 < MA10*(1-MARate):
        Amount += -50
        MAFlag[2] = -1
    return Amount,MAFlag

### 判断上次操作与本次操作是否相同。返回ActFlag为0表示操作相同，不能进行；ActFlag不为0，则可进行，1为开，-1为平
def CheckActFalg(LastAct,Amount):
    ActOKFlag = 0
    if Amount != 0:
        ThisAct = Amount / abs(Amount)
        if (LastAct == -3) and ThisAct == 1:    #上次为休整，这次为开
            ActOKFlag = 2
        elif (LastAct == 0 or LastAct == -1 or LastAct == -2) and ThisAct == 1:    #上次为初始或平或利润率100%后强平，这次为开
            ActOKFlag = 1
        elif LastAct == 1 and ThisAct == -1:    #上次为开，这次为平
            ActOKFlag = -1
    #print("Amount:%d,ThisAct:%d,LastAct:%d,ActOKFlag:%d" % (Amount, ThisAct, LastAct, ActOKFlag),end = ' ')
    return ActOKFlag

### 判断均值交叉情况。返回MAOKFlag为1，表示均值交叉
def CheckMAFlag(OriginMAFlag,MAFlag):
    MAOKFlag = 0
    if operator.eq(OriginMAFlag,[0,0,0]): #均值交叉标识为初始值时，
        MAOKFlag = 0
    elif not(operator.eq(OriginMAFlag,MAFlag)): #均值交叉标识与存储值不同时，表示发生均值交叉
        MAOKFlag = 1
    return MAOKFlag

### 计算实际交易数量。开时，最多用光所有资金；平时，最多全部平掉
def CheckAmount(ActFlag, Amount, LivePX, TodayMoney, TodayPos):
    ActualAmount = 0
    if ActFlag == 1:    #开
        MaxAmount = int(TodayMoney/LivePX)    #舍尾取整
        if MaxAmount == 0:
            print(u"没有足够的资金！")
        elif MaxAmount < Amount:
            ActualAmount = MaxAmount
        else:
            ActualAmount = Amount
    elif ActFlag == -1: #平
        if TodayPos < abs(Amount):
            ActualAmount = TodayPos
        else:
            ActualAmount = abs(Amount)
    return ActualAmount

### 判断是否进行交易。个条件，先后顺序：开平交替，MA均线交叉，当天次数最多5次
# 结束休整期后，不管均线是否交叉，MA计算结果可开就开
def CheckFlag(LastAct, Amount, OriginMAFlag, MAFlag, NumAct,LivePX, TodayMoney, TodayPos):
    ActFlag = 0
    ActualAmount = 0
    ActOKFlag = CheckActFalg(LastAct, Amount)   #判断本次与上次操作是否开平交替
    if ActOKFlag == 2:
        ActFlag = 1
        ActualAmount = CheckAmount(ActFlag, Amount, LivePX, TodayMoney, TodayPos)  # 计算真实应该交易的数量
        if ActualAmount == 0: return 0,0
    elif ActOKFlag != 0:
        MAOKFlag = CheckMAFlag(OriginMAFlag, MAFlag)  #判断MA均线是否交叉
        if (MAOKFlag == 1) and (NumAct <=4):    #均线交叉，且当天交易最多5次
            ActFlag = ActOKFlag
            ActualAmount = CheckAmount(ActFlag, Amount, LivePX, TodayMoney, TodayPos)  # 计算真实应该交易的数量
            if ActualAmount == 0: return 0,0
    return ActFlag, ActualAmount

### 测试交易（未接入Wind）
def Trade(ActFlag,ActualAmount,TaxRate,LivePX,TodayMoney, TodayPos):
    if ActFlag == 1:
        TodayPos += ActualAmount
        TodayMoney -= LivePX*ActualAmount*(1+TaxRate)
    elif ActFlag == -1:
        TodayPos -= ActualAmount
        TodayMoney += LivePX*ActualAmount*(1-TaxRate)
    return TodayMoney, TodayPos

### 模拟交易（接入Wind模拟交易平台）
def SimulateTrade(Kind,AccountInfo,ActFlag,ActualAmount,TaxRate,LivePX,TodayMoney,TodayPos):
    if ActFlag == 1:    #开
        data1 = w.tlogon(BrokerID=AccountInfo[0], DepartmentID=0, LogonAccount=AccountInfo[1], Password=AccountInfo[2], AccountType='shf')  #登录模拟账户
        data21 = w.torder(SecurityCode=Kind, TradeSide='Buy', OrderPrice=LivePX, OrderVolume=ActualAmount, MarketType='SHF')  #模拟开仓
        TodayPos += ActualAmount
        TodayMoney -= LivePX*ActualAmount*(1+TaxRate)
    elif ActFlag == -1:     #平
        data1 = w.tlogon(BrokerID=AccountInfo[0], DepartmentID=0, LogonAccount=AccountInfo[1], Password=AccountInfo[2], AccountType='shf')  #登录模拟账户
        data21 = w.torder(SecurityCode=Kind, TradeSide='Sell', OrderPrice=LivePX, OrderVolume=ActualAmount,MarketType='SHF')  #模拟平仓
        TodayPos -= ActualAmount
        TodayMoney += LivePX*ActualAmount*(1-TaxRate)
    #print(TodayMoney, TodayPos)
    return TodayMoney, TodayPos

### 检查利润是否达到100%
def CheckProfit(TodayMoney, TodayPos,LivePX,OriginAsset):
    NowAsset = TodayMoney+TodayPos*LivePX
    if NowAsset >= 2*OriginAsset: return True
    else: return False

### 计算收益
def calc_profit(TotalMoney, TotalPos, LastPrice, TodayMoney, TodayPos, closeprice, OriginAsset):
    money_delta = TodayMoney - TotalMoney
    pos_delta = TodayPos - TotalPos
    posvalue = (TodayPos*closeprice)-(TotalPos*LastPrice)
    profit_delta = money_delta + posvalue
    accumulatedProfit = TodayMoney+TodayPos*closeprice-OriginAsset
    profit = [TodayMoney, TodayPos, closeprice, money_delta, pos_delta, posvalue, profit_delta,accumulatedProfit]
    return profit


### 实时数据画图
def Plot_Live_MAProfit(LogFileName, DayFileName):
    PX, MA5, MA10, MA20 = np.loadtxt(fname=LogFileName, delimiter=',', skiprows=1, usecols=(1, 2, 3, 4), unpack=True)
    DateTime = np.loadtxt(fname=LogFileName, dtype=str, delimiter='_', skiprows=1, usecols=0, unpack=True)
    Money, Pos, totAsset, dailyassetdelta, assetdelta = np.loadtxt(fname=DayFileName, delimiter=',', skiprows=1,usecols=(1, 2, 4, 7, 8), unpack=True)
    RecDay = np.loadtxt(fname=DayFileName, dtype=str, delimiter=',', skiprows=1, usecols=0, unpack=True)
    plt.rcParams['figure.figsize'] = (20.0, 16.0)
    plt.rcParams['lines.linewidth'] = 1
    plt.subplot(2,2,1)
    plt.xlim(xmin=0, xmax=len(PX) - 1)
    plt.plot(PX, color='black', label='Live Price')
    plt.plot(MA5, color='red', label='5-day Mean')
    plt.plot(MA10, color='green', label='10-day Mean')
    plt.plot(MA20, color='blue', label='20-day Mean')
    locs1, labels1 = getLocLabels(DateTime, 10)
    plt.xticks(locs1,labels1,rotation=45)
    #plt.xticks(rotation=45)
    plt.legend()
    plt.title('Price and 5/10/20-day Mean')
    plt.subplot(2,2,2)
    if totAsset.size>1:
        plt.xlim(xmin=0, xmax=len(Pos) - 1)
        plt.plot(totAsset, color='black', label='Total Asset')
        locs2, labels2 = getLocLabels(RecDay, 10)
        plt.xticks(locs2, labels2,rotation=45)
    #plt.xticks(rotation=45)
    plt.title('Total Asset')
    plt.subplot(2,2,3)
    if dailyassetdelta.size>1:
        plt.xlim(xmin=0, xmax=len(Pos) - 1)
        plt.plot(dailyassetdelta, color='blue', label='Daily Asset Change')
        plt.plot(assetdelta, color='red', label='Accumulated Asset Change')
        plt.xticks(locs2, labels2,rotation=45)
    #plt.xticks(rotation=45)
    plt.legend()
    plt.title('Asset Change')
    plt.subplot(2,2,4)
    if Pos.size>1:
        plt.xlim(xmin=0, xmax=len(Pos) - 1)
        plt.plot(Pos, color='blue', label='Total Postion')
        plt.xticks(locs2, labels2,rotation=45)
    #plt.xticks(rotation=45)
    plt.title('Total Postion')
    plt.subplots_adjust(hspace=0.4)
    plt.savefig("test.png",dpi=300)
    plt.show()





def getMinuteData(HistoryData):
    MinutePX = HistoryData[3]
    #MinuteDay = HistoryData[1]
    MinuteTime = str(HistoryData[2])[0:2]+':'+str(HistoryData[2])[2:4]
    return MinuteTime,MinutePX


def MAplot2(MARecords):
    XXX = []
    YPX = []
    YMA5 = []
    YMA10 = []
    YMA20 = []
    for i in range(0, len(MARecords)):
        XXX.append(i + 1)
        YPX.append(MARecords[i][2])
        YMA5.append(MARecords[i][3])
        YMA10.append(MARecords[i][4])
        YMA20.append(MARecords[i][5])
    #plt.rcParams['savefig.dpi'] = 200
    #plt.rcParams['figure.dpi'] = 200
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.ion()
    for i in range(0,len(XXX)):
        plt.plot(XXX[i], YPX[i], color='black', label=u'成交价')
        plt.plot(XXX[i], YMA5[i], color='red', label=u'5日均值')
        plt.plot(XXX[i], YMA10[i], color='green', label=u'10日均值')
        plt.plot(XXX[i], YMA20[i], color='blue', label=u'20日均值')
        plt.pause(0.5)
        plt.show()
    plt.legend()
    #plt.savefig('History_LastPX&MA.png', dpi = 600)
    #plt.show()

def MAplot(MARecords):
    XXX = []
    YPX = []
    YMA5 = []
    YMA10 = []
    YMA20 = []
    for i in range(0, len(MARecords)):
        XXX.append(i + 1)
        YPX.append(MARecords[i][2])
        YMA5.append(MARecords[i][3])
        YMA10.append(MARecords[i][4])
        YMA20.append(MARecords[i][5])
    #plt.rcParams['savefig.dpi'] = 200
    #plt.rcParams['figure.dpi'] = 200
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.plot(XXX, YPX, color='black', label=u'成交价', linewidth=1)
    plt.plot(XXX, YMA5, color='red', label=u'5日均值', linewidth=1)
    plt.plot(XXX, YMA10, color='green', label=u'10日均值', linewidth=1)
    plt.plot(XXX, YMA20, color='blue', label=u'20日均值', linewidth=1)
    plt.legend()
    plt.savefig('History_LastPX&MA.png', dpi = 600)
    plt.show()

def ReadHistoryData(filename):
    file = xlrd.open_workbook(filename)
    Amount = 0
    oldrecord = {}
    for i in range(0,file.nsheets):
        sheet = file.sheet_by_index(i)
        for j in range(1,sheet.nrows):
            ContractID = sheet.cell_value(j,0)
            OriginT = xlrd.xldate_as_tuple(sheet.cell_value(j,1),0)
            RecDate = ("{}{:0>2}{:0>2}").format(int(OriginT[0]), int(OriginT[1]), int(OriginT[2]))
            RecTime = ("{:0>2}{:0>2}").format(int(OriginT[3]),int(OriginT[4]))
            LastPX = sheet.cell_value(j,2)
            oldrecord[Amount] = [ContractID,RecDate,RecTime,LastPX]
            Amount += 1
    return oldrecord

def getDayPX(HistoryData):
    HistoryDayPX = {}
    Amount = 0
    daylist = [HistoryData[0][1]]
    for i in range(0,len(HistoryData)):
        day = HistoryData[i][1]
        if not (day in daylist):
            if i < len(HistoryData)-1:
                HistoryDayPX[Amount] = [i-1,HistoryData[i-1][1],HistoryData[i-1][3]]
                Amount += 1
                daylist.append(day)
            elif i == len(HistoryData)-1:
                HistoryDayPX[Amount] = [i-1,HistoryData[i - 1][1], HistoryData[i - 1][3]]
                HistoryDayPX[Amount+1] = [i,HistoryData[i][1], HistoryData[i][3]]
        else:
            if i == len(HistoryData) - 1:
                HistoryDayPX[Amount] = [i,HistoryData[i][1], HistoryData[i][3]]
    return HistoryDayPX

def calc_HistoryPreMA(HistoryDayPX,index):
    PreMA20 = 0
    PreMA10 = 0
    PreMA5 = 0
    for i in range(index-19,index):
        PreMA20 += HistoryDayPX[i][2]
    for i in range(index-9,index):
        PreMA10 += HistoryDayPX[i][2]
    for i in range(index-4,index):
        PreMA5 += HistoryDayPX[i][2]
    return (PreMA5, PreMA10, PreMA20)

def calc_HistoryMA(PreMA5, PreMA10, PreMA20, HistoryMin1PX):
    MA5 = (PreMA5 + HistoryMin1PX) / 5
    MA10 = (PreMA10 + HistoryMin1PX) / 10
    MA20 = (PreMA20 + HistoryMin1PX) / 20
    return (MA5, MA10, MA20)

def HistoryTrade(LastAct, TaxRate, HistoryMin1PX, HistoryMin1Time, Amount, TodayMoney, TodayPos, Records):
    LivePrice = HistoryMin1PX
    livetime = HistoryMin1Time
    # 判断是否第一次操作。若否，则判断是否与上次操作相同
    actflag = -1
    ThisAct = Amount / abs(Amount)
    if len(Records) == 0:
        if LastAct != ThisAct:
            actflag = 1
            LastAct = ThisAct
    else:
        LastAct = Records[len(Records)-1][1] / abs(Records[len(Records)-1][1])
        if LastAct != ThisAct:
            actflag = 1
            LastAct = ThisAct
    # 买入操作，需判断是否有足够的资金
    if actflag == 1:
        if Amount > 0:
            temp = int(TodayMoney / LivePrice)
            if temp < Amount:
                RealAmount = temp
                #print(u"资金不足,", )
            else:
                RealAmount = Amount
            TodayPos += RealAmount
            TodayMoney -= RealAmount * LivePrice * (1 + TaxRate)
            record = (str(livetime), RealAmount, LivePrice, TodayMoney, TodayPos)
            Records.append(record)
            print(u"%s: 此次买入%d手，成交价%s元。当前资金%s元，共计持仓%d手。" % (livetime, RealAmount, format(LivePrice, '.0f'), format(TodayMoney, ',.0f'), TodayPos))
        # 卖出操作，需判断是否持有足够的手数
        elif Amount < 0:
            Amount = abs(Amount)
            if TodayPos < Amount:
                RealAmount = TodayPos
                #print(u"持仓不足, 全部卖出。", )
            else:
                RealAmount = Amount
            TodayPos -= RealAmount
            TodayMoney += RealAmount * LivePrice * (1 - TaxRate)
            record = (str(livetime), -1 * RealAmount, LivePrice, TodayMoney, TodayPos)
            Records.append(record)
            print(u"%s: 此次买出%d手，成交价%s元。当前资金%s元，共计持仓%d手。" % (livetime, RealAmount, format(LivePrice, '.0f'), format(TodayMoney, ',.0f'), TodayPos))
    return (TodayMoney, TodayPos, Records)


def PrintDayProfit(MinuteDay,DayProfit,DayFile):
    TodayMoney = DayProfit[0]
    TodayPos = DayProfit[1]
    TodayAsset = DayProfit[0]+DayProfit[1]*DayProfit[2]
    moneydelta = DayProfit[3]
    posdelta = DayProfit[4]
    posvalue = DayProfit[5]
    profitdelta = DayProfit[6]
    accumulatedprofit = DayProfit[7]
    print(u"日期:%s, 资金:%12.2f, 持仓:%4d, 总资产:%12.2f, 资金变化:%12.2f, 持仓变化:%4d, 持仓价值变化:%12.2f, 总资产变化(前一天):%12.2f, 总资产变化(累积):%12.2f" % (MinuteDay, TodayMoney, TodayPos, TodayAsset, moneydelta, posdelta, posvalue, profitdelta,accumulatedprofit))
    print("%s,%12.2f,%4d,%12.2f,%12.2f,%4d,%12.2f,%12.2f,%12.2f" % (MinuteDay, TodayMoney, TodayPos, TodayAsset, moneydelta, posdelta, posvalue, profitdelta,accumulatedprofit), file=DayFile)
    return DayFile

def PXMAplot(LogFileName):
    PX,MA5,MA10,MA20 = np.loadtxt(fname=LogFileName,delimiter=',',skiprows=1,usecols=(1,2,3,4),unpack=True)
    DateTime = np.loadtxt(fname=LogFileName,dtype=str,delimiter=',',skiprows=1,usecols=0,unpack=True)
    plt.plot(PX, color='black', label='Price')
    plt.plot(MA5, color='red', label='5-day Mean')
    plt.plot(MA10, color='green', label='10-day Mean')
    plt.plot(MA20, color='blue', label='20-day Mean')
    plt.legend()
    #plt.savefig("MA.svg")
    plt.savefig("MA.png",dpi=600)
    plt.show()

def DayProfitplot(DayFileName):
    Money,Pos,totAsset,dailyassetdelta,assetdelta= np.loadtxt(fname=DayFileName, delimiter=',', skiprows=1, usecols=(1,2,3,7,8), unpack=True)
    RecDay = np.loadtxt(fname=DayFileName, dtype=str, delimiter=',', skiprows=1, usecols=0, unpack=True)
    #plt.plot(totAsset, color='black', label='Total Asset')
    plt.plot(dailyassetdelta, color='blue', label='Daily Asset Change')
    plt.plot(assetdelta, color='red', label='Accumulated Asset Change')
    plt.legend()
    # plt.savefig("DayRecord.svg")
    plt.savefig("DayRecord.png", dpi=600)
    plt.show()

def getLocLabels(list,nums):
    locs = []
    labels = []
    if nums >= len(list):
        for i in range(0, len(list)):
            label = list[i].split('_')[0]
            locs.append(i)
            labels.append(label)
    else:
        interval = int(len(list)/(nums-1))
        i = 0
        while i<len(list):
            loc = i
            label = list[i].split('_')[0]
            locs.append(loc)
            labels.append(label)
            i += interval
        if (i-len(list)) <= (interval*2/3):
            locs.append(len(list)-1)
            labels.append(list[-1].split('_')[0])
    return locs,labels



def Plot_MA_Profit(LogFileName,DayFileName):
    PX, MA5, MA10, MA20 = np.loadtxt(fname=LogFileName, delimiter=',', skiprows=1, usecols=(1,2,3,4), unpack=True)
    DateTime = np.loadtxt(fname=LogFileName, dtype=str, delimiter=',', skiprows=1, usecols=0, unpack=True)
    Money, Pos, totAsset, dailyassetdelta, assetdelta = np.loadtxt(fname=DayFileName, delimiter=',', skiprows=1,usecols=(1,2,3,7,8), unpack=True)
    RecDay = np.loadtxt(fname=DayFileName, dtype=str, delimiter=',', skiprows=1, usecols=0, unpack=True)
    plt.rcParams['figure.figsize'] = (20.0, 16.0)
    plt.rcParams['lines.linewidth'] = 1
    locs1, labels1 = getLocLabels(DateTime, 10)
    locs2, labels2 = getLocLabels(RecDay, 10)
    plt.subplot(2,2,1)
    plt.xlim(xmin=0, xmax=len(PX) - 1)
    plt.plot(PX, color='black', label='Price')
    plt.plot(MA5, color='red', label='5-day Mean')
    plt.plot(MA10, color='green', label='10-day Mean')
    plt.plot(MA20, color='blue', label='20-day Mean')
    plt.xticks(locs1,labels1,rotation=45)
    plt.legend()
    plt.title('Price and 5/10/20-day Mean')
    plt.subplot(2,2,2)
    plt.xlim(xmin=0, xmax=len(Pos) - 1)
    plt.plot(totAsset, color='black', label='Total Asset')
    plt.xticks(locs2, labels2,rotation=45)
    plt.title('Total Asset')
    plt.subplot(2,2,3)
    plt.xlim(xmin=0, xmax=len(Pos) - 1)
    plt.plot(dailyassetdelta, color='blue', label='Daily Asset Change')
    plt.plot(assetdelta, color='red', label='Accumulated Asset Change')
    plt.xticks(locs2, labels2,rotation=45)
    plt.legend()
    plt.title('Asset Change')
    plt.subplot(2,2,4)
    plt.xlim(xmin=0, xmax=len(Pos) - 1)
    plt.plot(Pos, color='blue', label='Total Postion')
    plt.xticks(locs2, labels2,rotation=45)
    plt.title('Total Postion')
    plt.subplots_adjust(hspace=0.4)
    plt.savefig("MA&Profit.png",dpi=300)
    plt.show()

def readOldData():
    thisyear = datetime.now().year
    thismonth = datetime.now().month
    thisday = datetime.now().day
    quarterindex = int(thismonth / 3.3) + 1
    if thisday == 1 and thismonth == 1:
        oldfilename = "Record-%sQ%s.xlsx" % (thisyear - 1, 4)
    elif thisday == 1 and (thismonth == 4 or thismonth == 7 or thismonth == 10):
        oldfilename = "Record-%sQ%s.xlsx" % (thisyear, quarterindex - 1)
    else:
        oldfilename = "Record-%sQ%s.xlsx" % (thisyear, quarterindex)
    if os.path.isfile(oldfilename):
        oldfile = xlrd.open_workbook(oldfilename)
        oldsheet = oldfile.sheets()[0]
        TotalMoney = oldsheet.cell(0, 1).value
        TotalPos = oldsheet.cell(1, 1).value
        LastPrice = oldsheet.cell(2, 1).value
        OldRecords = []
        oldcount = 0
        if oldsheet.col_values(4)[1]:
            oldcount = int(oldsheet.col_values(4)[1])
        rectime = ''
        for i in range(0, oldcount):
            linedata = oldsheet.row_values(i+1)[5:]
            if oldsheet.cell(i+1,5).ctype == 1:
                rectime = oldsheet.cell(i + 1, 5).value
            elif oldsheet.cell(i+1,5).ctype == 3:
                temptime = xlrd.xldate_as_tuple(oldsheet.cell(i+1,5).value,0)
                rectime = datetime(*temptime)
            oldrecord = (rectime, linedata[1], linedata[2], linedata[3], linedata[4])
            OldRecords.append(oldrecord)
        lastact = OldRecords[0][1] / abs(OldRecords[0][1])
        return (TotalMoney, TotalPos, LastPrice, OldRecords, lastact)
    else:
        print (u"前一交易日的记录文件%s没有找到。" % oldfilename)



def buyORsell(lastact, TaxRate, livedata, Amount, TodayMoney, TodayPos, Records):
    LivePrice = livedata.Data[0][0]
    #判断是否第一次操作。若否，则判断是否与上次操作相同
    actflag = -1
    ThisAct = Amount / abs(Amount)
    if len(Records) == 0:
        if lastact != ThisAct:
            actflag = 1
            lastact = ThisAct
    else:
        lastact = Records[len(Records)-1][1] / abs(Records[len(Records)-1][1])
        if lastact != ThisAct:
            actflag = 1
            lastact = ThisAct
    # 买入操作，需判断是否有足够的资金
    if actflag == 1:
        if Amount > 0:
            temp = int(TodayMoney / LivePrice)
            if temp < Amount:
                RealAmount = temp
                print (u"资金不足,",)
            else:
                RealAmount = Amount
            TodayPos += RealAmount
            TodayMoney -= RealAmount * LivePrice * (1 + TaxRate)
            record = (str(datetime.now()).split('.')[0], RealAmount, LivePrice, TodayMoney, TodayPos)
            Records.append(record)
            print (u"此次买入%d手，成交价%s元。当前资金%s元，共计持仓%d手。" % (RealAmount, format(LivePrice,'.0f'), format(TodayMoney,',.0f'), TodayPos))
        # 卖出操作，需判断是否持有足够的手数
        elif Amount < 0:
            Amount = abs(Amount)
            if TodayPos < Amount:
                RealAmount = TodayPos
                print (u"持仓不足, 全部卖出。",)
            else:
                RealAmount = Amount
            TodayPos -= RealAmount
            TodayMoney += RealAmount * LivePrice  * (1 - TaxRate)
            record = (str(datetime.now()).split('.')[0], -1*RealAmount, LivePrice, TodayMoney, TodayPos)
            Records.append(record)
            print (u"此次买出%d手，成交价%s元。当前资金%s元，共计持仓%d手。" % (RealAmount, format(LivePrice,'.0f'), format(TodayMoney,',.0f'), TodayPos))
    return (TodayMoney, TodayPos, Records)

def printprofit(profit):
    monstr = printstatus(profit[0])
    posstr = printstatus(profit[1])
    valstr = printstatus(profit[2])
    totstr = printstatus(profit[3])
    print (u"现有资金%s，持仓数%s，持仓价值%s，总资产%s。" % (monstr, posstr, valstr, totstr))

def writeRecord(ResFile, TodayMoney, TodayPos, closeprice, profit, Records, OldRecords):
    sheet1 = ResFile.add_worksheet()
    #输出格式
    sheet1.set_column('A:A', 16)
    sheet1.set_column('B:B', 12)
    sheet1.set_column('E:E', 6)
    sheet1.set_column('F:F', 21)
    sheet1.set_column('G:J', 9)
    sheet1.set_column('I:I', 12)
    summary_format = ResFile.add_format({'bold': True, 'bg_color':'#FFFF66'})
    summarymoney_format = ResFile.add_format({'num_format': '#,##0','bg_color':'#FFFFCC'})
    delta_format = ResFile.add_format({'bold': True, 'bg_color': '#CCFFFF'})
    title_format = ResFile.add_format({'bold': True,'align':'center'})
    money_format = ResFile.add_format({'num_format':'#,##0'})
    align_format = ResFile.add_format({'align':'center'})
    #输出当天摘要情况
    summarytitle = [u"现有资金", u"持仓数", u"持仓单价", u"持仓总价值", u"总资产"]
    summary = [TodayMoney, TodayPos, closeprice, TodayPos*closeprice, TodayMoney+TodayPos*closeprice]
    for i in range(0, len(summarytitle)):
        sheet1.write(i, 0, summarytitle[i],summary_format)
        sheet1.write(i, 1, summary[i],summarymoney_format)
    profittitle = [u"资金变化", u"持仓数变化", u"持仓总价值变化", u"总资产变化"]
    temp = str(date.today())+u"变化情况："
    sheet1.merge_range(len(summarytitle)+1, 0, len(summarytitle)+1, 1, temp, title_format)
    #sheet1.write(len(summarytitle)+1, 0, temp, title_format)
    for i in range(0, len(profittitle)):
        sheet1.write(i+len(summarytitle)+2, 0, profittitle[i],delta_format)
        sheet1.write(i+len(summarytitle)+2, 1, profit[i],money_format)
    #输出当天交易明细
    Records.reverse()
    Records.extend(OldRecords)
    sheet1.write(0, 3, u"交易记录：",title_format)
    recordtitle = [u"序号", u"日期&时间", u"交易手数", u"成交价", u"可用资金", u"现持仓数"]
    for i in range(0, len(recordtitle)):
        sheet1.write(0, i+4, recordtitle[i],title_format)
    for i in range(0, len(Records)):
        sheet1.write(i + 1, 4, len(Records) - i,align_format)
        sheet1.write(i + 1, 5, Records[i][0],align_format)
        sheet1.write(i + 1, 6, Records[i][1],money_format)
        sheet1.write(i + 1, 7, Records[i][2],money_format)
        sheet1.write(i + 1, 8, Records[i][3],money_format)
        sheet1.write(i + 1, 9, Records[i][4],money_format)
    return ResFile