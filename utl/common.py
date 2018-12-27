#-*- coding:gb18030 -*-

from math import floor
import datetime

#数值取两位小数点精度
def toNumberFmt(value):
    if value is None:
        return 0
    return floor(value * 100 + 0.50001) / 100

def getLastDay(stlmDate):
    date = datetime.datetime(int(stlmDate[:4]), int(stlmDate[4:6]), int(stlmDate[6:8])) - datetime.timedelta(days=1)
    time_format = date.strftime('%Y%m%d')
    return time_format

def getNextDay(stlmDate):
    date = datetime.datetime(int(stlmDate[:4]), int(stlmDate[4:6]), int(stlmDate[6:8])) + datetime.timedelta(days=1)
    time_format = date.strftime('%Y%m%d')
    return time_format