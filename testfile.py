# coding=utf-8

##from datetime import datetime, date, timedelta
##TodayDate = str(datetime.date(datetime.now()))
##print ("%s Records of %s %s" % ('-'*10,TodayDate,'-'*10))
from functions import *
TodayDate = str(datetime.date(datetime.now()))
print(TodayDate)
##LogFileName = getDetailFileName()
##DayFileName = "Day_Close.txt"
##Plot_Live_MAProfit(LogFileName, DayFileName)
#
#timelist = ["12:25","12:27","12:28", "12:29"]
#
##print(nowtime)
##for i in range(0,len(timelist)):
##    tt = timelist[i].split(':')
##    tlimit.append(nowtime.replace(hour=int(tt[0]),minute=int(tt[1]),second=0))
##print(tlimit)
##for i in range(0,len(tlimit)):
##    if ((i+1) <= (len(tlimit)-1)) and (tlimit[i]< nowtime) and  (nowtime< tlimit[i+1]):
##        print(i,'Yes')
#
#def TransTimeList(timelist):
#    nowtime = datetime.now()
#    timelist2 = []
#    if (len(timelist)%2) != 0:
#        print(u"时间范围首尾不配对！")
#        exit()
#    for i in range(0, len(timelist)):
#        tt = timelist[i].split(':')
#        timelist2.append(nowtime.replace(hour=int(tt[0]), minute=int(tt[1]), second=0))
#    for i in range(0,len(timelist2)-1):
#        if not (timelist2[i] < timelist2[i+1]):
#            print(u"时间先后顺序不对",timelist2[i],timelist2[i+1])
#            exit()
#    return timelist2
#
#def CheckTime(timelist2):
#    nowtime = datetime.now()
#    if nowtime > timelist2[-1]:
#        print(nowtime,timelist2[-1])
#        return False
#    else:
#        return True
#    #timelist2 = []
#    #if (len(timelist)%2) != 0:
#    #    print(u"时间范围首尾不配对！")
#    #    exit()
#    #for i in range(0, len(timelist)):
#    #    tt = timelist[i].split(':')
#    #    timelist2.append(nowtime.replace(hour=int(tt[0]), minute=int(tt[1]), second=0))
#    #interval = 1
#    ##for i in range(0, int(len(timelist2)/2)):
#    ##    if (timelist2[2*i]< nowtime) and (nowtime< timelist2[2*i+1]):
#    ##        flag = 1
#    #if nowtime < timelist2[0]:
#    #    print(timelist2[0],nowtime)
#    #    interval = (timelist2[0]-nowtime).seconds
#    #    return interval
#    #for i in range(0, int(len(timelist2)/2)):
#    #    if (timelist2[2*i]< nowtime) and (nowtime< timelist2[2*i+1]):
#    #        interval = 1
#    #        return interval
#    #for i in range(0,int(len(timelist2)/2)):
#    #    if nowtime < timelist2[2*i]:
#    #        interval = (timelist2[2*i]-nowtime).seconds
#    #        return interval
#def CheckInterval(timelist2):
#    interval = 1
#    nowtime = datetime.now()
#    if nowtime < timelist2[0]:
#        print(nowtime,timelist2[0])
#        interval = (timelist2[0]-nowtime).seconds
#    else:
#        for i in range(0, int(len(timelist2)/2)):
#            if (timelist2[2*i] <= nowtime) and (nowtime <= timelist2[2*i+1]):
#                interval = 1
#                break
#            elif (timelist2[2*i+1] < nowtime) and (nowtime < timelist2[2*i+2]):
#                interval = max((timelist2[2*i+2]-nowtime).seconds,1)
#                break
#    return interval
#
#timelist2 = TransTimeList(timelist)
#while CheckTime(timelist2):
#    interval = CheckInterval(timelist2)
#    print(interval,datetime.now())
#    time.sleep(interval)
