#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#推送信息

import cx_Oracle
from multiprocessing import Pool
import os
from hashlib import sha1
from datetime import datetime
import time
import random
import string
import requests

appSecret = '2fec1d6e71bc'
appKey="9a0183cea3f2f0e5df646017661bfcb7";
proxy = {}

def getCheckSum(appSecret, nonce, curTime):
    signdata = appSecret + nonce + curTime
    s1 = sha1()
    s1.update(signdata.encode("utf-8"))
    return s1.hexdigest()


def get_header():
    now = datetime.now()
    curTime = str(int(now.timestamp()))
    nonce = ''.join(random.sample(string.ascii_letters + string.digits, 20))
    checksum = getCheckSum(appSecret, nonce, curTime)
    Content_Type = "application/x-www-form-urlencoded;charset=utf-8"
    header = {'Content-Type': Content_Type, 'AppKey': appKey, 'Nonce': str(nonce), 'CurTime': curTime,'CheckSum': checksum}
    return header

def get_body(toUser, msg):
    data = 'from=zhangsan&ope=0&to=%s&type=0&body={"msg":"%s"}' % (toUser, msg)
    return data.encode('utf-8')

def sendMsg(toUser, msg):
    url = 'https://api.netease.im/nimserver/msg/sendMsg.action'
    postdata = get_body(toUser, msg)
    head = get_header()
    response = requests.post(url, data=postdata, headers=head, proxies=proxy)
    return response.json()

def notice_agent(db, pl):


def main():
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),
                           encoding='gb18030')
    pl = Pool(10)
    #获取信息

    while 1:
        notice_agent(db, pl)
        time.sleep(1)


if __name__ == '__main__':
    main()


