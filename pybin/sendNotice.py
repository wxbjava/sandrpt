#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#推送信息

import cx_Oracle
from multiprocessing import Pool
import os
from hashlib import sha1
import time
import random
import string
import requests
from utl.common import *
import logging as log

log.basicConfig(level=log.INFO,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                datefmt='%a, %d %b %Y %H:%M:%S',
                handlers={log.FileHandler(filename=os.environ['HOME'] + '/log/sendNotice.log', mode='a', encoding='gb18030')})

appSecret = os.environ['APPSECRET']
appKey = os.environ['APPKEY']
proxy = {
    'http':'172.17.2.19:3128',
    'https':'172.17.2.19:3128'
}
proxy = {}  #开发环境不配置代理

def getCheckSum(appSecret, nonce, curTime):
    signdata = appSecret + nonce + curTime
    s1 = sha1()
    s1.update(signdata.encode("utf-8"))
    return s1.hexdigest()


def get_header():
    curTime = getCurrUtcStr()
    nonce = ''.join(random.sample(string.ascii_letters + string.digits, 20))
    checksum = getCheckSum(appSecret, nonce, curTime)
    Content_Type = "application/x-www-form-urlencoded;charset=utf-8"
    header = {'Content-Type': Content_Type, 'AppKey': appKey, 'Nonce': str(nonce), 'CurTime': curTime,'CheckSum': checksum}
    return header

def get_body(toUser, msg):
    data = 'from=zhangsan&ope=0&to=%s&type=0&body={"msg":"%s"}' % (toUser, msg)
    log.info(data)
    return data.encode('utf-8')

def sendMsg(toUser, msg):
    url = 'https://api.netease.im/nimserver/msg/sendMsg.action'
    postdata = get_body(toUser, msg)
    head = get_header()
    response = requests.post(url, data=postdata, headers=head, proxies=proxy)
    log.info(response.json())

def updMsgSta(db, msgId):
    sql = "update TBL_SND_NOTICE_LOG set sta ='1' where ind = %d" % msgId
    cursor = db.cursor()
    cursor.execute(sql)
    db.commit()
    cursor.close()


def notice_agent(db, pl):
    end_time = getDayTime()
    start_time = getDayTime(diffSec=-300)
    sql = "select IND,ACCT_ID,MSG_CONTENT from TBL_SND_NOTICE_LOG where SEND_TIME <= '%s' and SEND_TIME >='%s' " \
          "and ACCT_TYPE ='01' and send_tp ='1' and sta ='0'" % (end_time, start_time)
    cursor = db.cursor()
    cursor.execute(sql)
    for ltData in cursor:
        #更新数据
        updMsgSta(db, ltData[0])
        pl.apply_async(sendMsg, args=(ltData[1], ltData[2]))
    cursor.close()


def reConnectDb(dbuser, dbpwd, tnsname):
    while True:
        try:
            db = cx_Oracle.connect('%s/%s@%s' % (dbuser, dbpwd, tnsname), encoding='gb18030')
            break
        except Exception as e:
            log.info(e)
            time.sleep(10)
    log.info('reconn success')
    return db

def work(n):
    log.info('test %d', n)

def main():
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['ONLDBUSER'], os.environ['ONLDBPWD'], os.environ['TNSNAME']),
                           encoding='gb18030')
    pl = Pool(10)
    #获取信息
    while 1:
        try:
            notice_agent(db, pl)
        except cx_Oracle.OperationalError as e:
            error, = e.args
            if error.code == 3113:
                log.info('db connect restart')
                db = reConnectDb(os.environ['ONLDBUSER'], os.environ['ONLDBPWD'], os.environ['TNSNAME'])
        time.sleep(1)


if __name__ == '__main__':
    main()


