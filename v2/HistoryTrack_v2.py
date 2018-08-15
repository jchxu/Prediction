# coding=utf-8


### 用户自定义设置 ###
TaxRate = 0.0007   #交易手续费率，万分之七
MARate = 0.01   #MA计算的浮动比例
OriginAsset = 1000000   #初始资金
TimeList = ["8:00","23:59"]   #每天交易的时间段。开始，结束……，成对列出
TimeInterval = 5   #程序运行的时间间隔，单位为秒
PlatformCode = '0000'   #交易/模拟平台（经纪商）代码
Account = 'W109839501202'   #交易/模拟账户名
Passwd = '000000'   #交易/模拟账户名密码
Kind = 'RB1901.SHF' #交易的商品品种
MonthList = [4,9,12]    #每年哪几个月的第二个星期五需要强制平仓
Contracts = ['RB1805','RB1810']     #指定交易品种

