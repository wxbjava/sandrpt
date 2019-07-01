#-*- coding:gb18030 -*-

from math import floor
import datetime

#��ֵȡ��λС���㾫��
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

def getDayTime(diffSec=0):
    now = datetime.datetime.now()
    date = now + datetime.timedelta(seconds=diffSec)
    time_format = date.strftime('%Y%m%d%H%M%S')
    return time_format

def getCurrUtcStr():
    now = datetime.datetime.now()
    return str(int(now.timestamp()))



#���ڼ���
def isHoliDay(db, stlmDate):
    cursor = db.cursor()
    sql = "select * from TBL_HOLI_INF where START_DATE <='%s' and END_DATE >'%s'" % (stlmDate, stlmDate)
    cursor.execute(sql)
    x = cursor.fetchone()
    if x is not None:
        cursor.close()
        return True
    cursor.close()
    return False

