#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#代理商推送信息

import cx_Oracle
from multiprocessing.pool import ThreadPool
import os
import time
from utl.neteim import NeteIm
from utl.common import *
import logging as log
import utl.proxycfg as proxycfg

log.basicConfig(level=log.INFO,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                datefmt='%a, %d %b %Y %H:%M:%S',
                handlers={log.FileHandler(filename=os.environ['HOME'] + '/log/sendNoticeAgent.log', mode='a', encoding='gb18030')})

appSecret = os.environ['APPSECRET']
appKey = os.environ['APPKEY']
proxy = proxycfg.get_proxy()


def sendMsg(toUser, msg):
    nete_im = NeteIm(appSecret, appKey, '10000', proxy)
    res = nete_im.send_msg(toUser, msg)
    for msg in res:
        log.info(msg)

def sendBatMsg(toUser, msg):
    nete_im = NeteIm(appSecret, appKey, '10000', proxy)
    res = nete_im.send_bat_msg(toUser, msg)
    for msg in res:
        log.info(msg)

def sendBatMsgByType(batType, msg):
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['ONLDBUSER'], os.environ['ONLDBPWD'], os.environ['TNSNAME']),
                           encoding='gb18030')
    #获取推送列表
    if batType == 'A0000001':
        #所有一代推送
        sql = "select agent_cd from tbl_mcht_agent where ext5='1' and EXT2='09970000'"
    elif batType == 'A0000002':
        #推送所有代理
        sql = "select agent_cd from tbl_mcht_agent where EXT2='09970000'"
    else:
        log.error("未知类型")
        return
    log.info(sql)
    cursor = db.cursor()
    cursor.execute(sql)
    i = 0
    toAccts = []
    for ltData in cursor:
        toAccts.append(ltData[0])
        i = i + 1
        if i == 500:
            sendBatMsg(toAccts, msg)
            toAccts = []
            i = 0
            time.sleep(0.6)

    if i > 0:
        sendBatMsg(toAccts, msg)
    cursor.close()
    db.close()

def updMsgSta(db, msgId):
    sql = "update TBL_SND_NOTICE_LOG set sta ='1' where ind = %d" % msgId
    cursor = db.cursor()
    cursor.execute(sql)
    db.commit()
    cursor.close()


def notice_agent(db, pl):
    end_time = getDayTime()
    start_time = getDayTime(diffSec=-300)
    sql = "select IND,trim(ACCT_ID),ACCT_TYPE,MSG_CONTENT from TBL_SND_NOTICE_LOG where SEND_TIME <= '%s' and SEND_TIME >='%s' " \
          "and ACCT_TYPE in ('01','03') and send_tp ='1' and sta ='0'" % (end_time, start_time)
    cursor = db.cursor()
    cursor.execute(sql)
    for ltData in cursor:
        #更新数据
        updMsgSta(db, ltData[0])
        if ltData[2] == '03':
            #类型批量发送
            pl.apply_async(sendBatMsgByType, args=(ltData[1], ltData[3],))
        else:
            pl.apply_async(sendMsg, args=(ltData[1], ltData[3],))
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


def main():
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['ONLDBUSER'], os.environ['ONLDBPWD'], os.environ['TNSNAME']),
                           encoding='gb18030')
    pl = ThreadPool(10)
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



