# coding=utf-8
from WindPy import w
data=w.start()
#data=w.wsd("600000.SH","close","2018-07-04")
print(data.ErrorCode)

#Kinds = ['RB1808.SHF','RB1901.SHF']  #交易的商品品种
#Kind = 'RB1901.SHF'
#for i in range(0,len(Kinds)):
#    LivePX = w.wsd(Kinds[i], "close", "ED-19TD", "2018-07-31", "Fill=Previous")
#    print(Kinds[i],":",LivePX)

#BrokerID = '0000'
#Account = 'W109839501202'
#Password = '000000'
#OrderPrice = 4200
#OrderVolume = 10
#
#### 登录（模拟）账号：
#data1 = w.tlogon(BrokerID=BrokerID,DepartmentID=0,LogonAccount=Account,Password=Password,AccountType='shf')  #经纪商代码，营业部代码，资金账号，密码，账号类型(上海商品)
#LogonID = data1.Data[0][0]
#
#### 委托交易：
#data21 = w.torder(SecurityCode='RB.SHF', TradeSide='Buy', OrderPrice=OrderPrice, OrderVolume=OrderVolume, MarketType='SHF')  #期货Wind代码，Buy/Sell/Short/Cover/CoverToday/SellToday，委托价格，交易数量
#RequestID = data21.Data[0][0]
#data22 = w.tquery(qrycode='Order',RequestID=RequestID)
#TradeSide = data22.Data[4][0]
#OrderPrice = data22.Data[5][0]
#OrderVolume = data22.Data[6][0]
#OrderTime = data22.Data[7][0]
##print(TradeSide,OrderPrice,OrderVolume,OrderTime)
#
#### 查询：
#data31 = w.tquery(qrycode=0)  #查询资金
#AvailableFund = data31.Data[1][0]
#BalanceFund = data31.Data[3][0]
##print(AvailableFund,BalanceFund)
#
#data32 = w.tquery(qrycode=1,WindCode='RB.SHF')  #查询持仓
#EnableVolume = data32.Data[6][0]
#TodayRealVolume = data32.Data[7][0]
#TodayOpenVolume = data32.Data[8][0]
##print(EnableVolume,TodayRealVolume,TodayOpenVolume)
#
##for i in range(0,len(data32.Fields)):
##    print(data32.Fields[i])
#
#### (模拟)账号登出：
#data4 = w.tlogout(LogonID)
#w.stop()