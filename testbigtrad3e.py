# coding=utf-8

from WindPy import w
from functions import *
import time


w.start()
timelist = ["21:00","23:00"]
timeinterval = 1
timelist2 = TransTimeList(timelist)
while CheckTime(timelist2):
    interval = CheckInterval(timeinterval, timelist2)
    data = w.wsq("RB.SHF", "rt_date,rt_time,rt_last,rt_latest,rt_last_vol,rt_oi_change,rt_nature")
    print(data.Data)
    time.sleep(interval)
w.stop()