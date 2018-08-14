# coding=utf-8
from WindPy import w
from datetime import datetime,timedelta
#from xlsxwriter import Workbook
#from functions import *
#import time
#import random
OriginMAFlag = [0,0,0]
print(OriginMAFlag)



#### 用户自定义设置 ###
#Records = []         #当天交易记录
#rate = 0.0007        #交易手续费率，万分之七
#timeinterval = 1     #每隔1分钟运行一次，单位为秒
######################
#
#TotalMoney, TotalPos, LastPrice, OldRecords, lastact = readOldData()   #获取昨日收盘后信息：可用资金，持仓数，收盘价
#w.start()   #打开Wind-Python接口
#olddata = w.wsd("RB.SHF", "close", "ED-20TD", datetime.today(), "Fill=Previous")    #获取前20个交易日的收盘价
#
#print (u"现有资金%s元，持仓%d手，持仓价值%s元，总资产%s元。" % (format(TotalMoney, ',.0f'), TotalPos, format(TotalPos*LastPrice,',.0f'), format(TotalMoney+TotalPos*LastPrice,',.0f')))
#print ('-'*9+" Today's Records "+'-'*9)   #分割线
#
#TodayMoney = TotalMoney
#TodayPos = TotalPos
#for i in range(10): #测试运行n次
##while str(datetime.now()) < (str(date.today())+' '+"14:00:00"): #假设当天下午5点停盘
#    livedata = w.wsq("RB.SHF", "rt_latest") #获取最新成交价
#    MA5, MA10, MA20 = calc_MA(olddata, livedata)    #计算MA5，MA10，MA20数据
#    #amount = judge_status(MA5, MA10, MA20)   #根据MA判断买入（amount为正），卖出(amount为负)，绝对值为开平仓手数
#    amount = random.sample([-80,-50,-20,30,80,120], 1)[0]    # 随机测试
#    TodayMoney, TodayPos, Records = buyORsell(lastact, rate, livedata, amount, TodayMoney, TodayPos, Records)
#    time.sleep(timeinterval)  #每隔n秒执行一遍
#closeprice = w.wsd("RB.SHF", "close", "ED0TD", date.today(), "Fill=Previous").Data[0][0]  # 获取获取今日收盘价
#profit = calc_profit(TotalMoney, TotalPos, LastPrice, TodayMoney, TodayPos, closeprice)     #计算今日资金/持仓变化
#
#print ('-'*9+" Today's Summary "+'-'*9)   #分割线
#print (u"现有资金%s元，持仓%d手，持仓价值%s元，总资产%s元。" % (format(TodayMoney, ',.0f'), TodayPos, format(TodayPos*closeprice,',.0f'), format(TodayMoney+TodayPos*closeprice,',.0f')))
#printprofit(profit)
#TotalMoney = TodayMoney
#TotalPos = TodayPos
#newfilename = "Record-%dQ%d.xlsx" % (datetime.now().year, int(datetime.now().month/3.3)+1)
#ResFile = Workbook(newfilename)
#writeRecord(ResFile, TodayMoney, TodayPos, closeprice, profit, Records, OldRecords)
#ResFile.close()

#test = [1,1,1,1,1,1,1,1,1,1,1,1,1]
#print(len(test))
#print(sum(test[len(test)-4:len(test)]))
#w.start()
#Kind = 'RB1901.SHF'
#olddata = w.wsd(Kind, "close", "ED-19TD", datetime.today(), "Fill=Previous")
#print(len(olddata.Times))
#for i in olddata.Times[15:19]:
#    print(i)
#import calendar
#TodayDate = str(datetime.date(datetime.now()))

#monthlist = [4,9,12]

#Fridays = []
#ThisYear = int(datetime.date(datetime.now()).year)
#for i in range(0,len(monthlist)):
#    num = 0
#    for j in range(1,calendar.monthrange(ThisYear,monthlist[i])[1]+1):
#        print(monthlist[i],j)
#        if datetime(year=ThisYear,month=monthlist[i],day=j).weekday() == calendar.FRIDAY:
#            if num == 1:
#                print(datetime(year=ThisYear,month=monthlist[i],day=j).strftime('%Y-%m-%d'))
#                Fridays.append(datetime(year=ThisYear,month=monthlist[i],day=j).strftime('%Y-%m-%d'))
#                break
#            else:
#                num += 1
#
#print(Fridays)

#def getFridays(monthlist):
#    Fridays = []
#    ThisYear = int(datetime.date(datetime.now()).year)
#    for i in range(0, len(monthlist)):
#        num = 0
#        for j in range(1, calendar.monthrange(ThisYear, monthlist[i])[1] + 1):
#            if datetime(year=ThisYear, month=monthlist[i], day=j).weekday() == calendar.FRIDAY:
#                if num == 1:
#                    Fridays.append(datetime(year=ThisYear, month=monthlist[i], day=j).strftime('%Y-%m-%d'))
#                    break
#                else:
#                    num += 1
#    return Fridays
#
#def checkRest(TodayDate,Fridays):
#    flag = 0
#    TodayList = TodayDate.split('-')
#    Today = datetime(year=int(TodayList[0]),month=int(TodayList[1]),day=int(TodayList[2]))
#    for i in range(0,len(Fridays)):
#        FridayList = Fridays[i].split('-')
#        Friday = datetime(year=int(FridayList[0]), month=int(FridayList[1]), day=int(FridayList[2]))
#        if (Today > Friday) and (Today <= (Friday+timedelta(days=30))):
#            flag = 1
#            break
#        else:
#            flag = 0
#    if flag == 1:
#        return True
#    elif flag == 0:
#        return False
#
#
#Fridays = getFridays(monthlist)
#print(Fridays)
##TodayDate = str(datetime.date(datetime.now()))  #获取当天时间，格式为yyyy-mm-dd
#TodayDate = "2018-09-15"
#if TodayDate in Fridays:
#    print(u"今日强制平仓！")
#elif checkRest(TodayDate,Fridays):
#    print(u"尚在强制平仓后的休整期！")
#
##dayOfWeek = datetime.today().weekday()
##print(dayOfWeek+1)