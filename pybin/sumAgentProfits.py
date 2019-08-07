#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#代理商分润统计

import cx_Oracle
import sys
import os
from utl.common import *

def insertAgentSum(db, stlmDate, agent_cd, txn_count, trans_amt,
                   trans_fee, all_profits, self_profits,
                   company_cd, daily_profits):
    cursor = db.cursor()
    #获取一代和层级
    sql = "select substrb(ext7,1,6), trim(ext5) from tbl_mcht_agent where agent_cd ='%s'" % agent_cd
    cursor.execute(sql)
    x = cursor.fetchone()
    if x is None:
        print("tbl_mcht_agent select error %s" % agent_cd)
        cursor.close()
        return False
    agent_1_cd = x[0]
    agent_level = x[1]

    date = getDayTime()
    sql = "INSERT INTO TBL_AGENT_SHARE_SUM (INST_DATE, INST_TIME, " \
          "AGENT_CD, HOST_DATE, TXN_COUNT, TRANS_AMT, TRANS_FEE, " \
          "ALL_PROFITS, SELF_PROFITS, CHARGE_FLAG, DAILY_PROFITS, " \
          "AGENT_1_CD, AGENT_LEVEL,  COMPANY_CD) VALUES " \
          "('%s', '%s', :1, %s, :2, :3, :4, :5, :6, :7, :8, :9, :10, :11)" % (date[0:8], date[8:14], stlmDate)
    cursor.prepare(sql)
    param = (agent_cd, txn_count, trans_amt, trans_fee,
             all_profits, self_profits, '1', daily_profits, agent_1_cd,
             agent_level, company_cd)
    cursor.execute(None, param)
    cursor.close()
    return True

def calcAgentProfitsSum(db, stlmDate):
    sql = "select agent_cd, count(1), sum(trans_amt), sum(trans_fee), sum(all_profits), " \
          "sum(self_profits),company_cd,sum(stlm_profits) from tbl_agent_share_dtl " \
          "where host_date = '%s'  group by agent_cd, company_cd" % stlmDate
    cursor = db.cursor()
    cursor.execute(sql)
    for ltData in cursor:
        agent_cd = ltData[0]
        txn_count = ltData[1]
        trans_amt = ltData[2]
        trans_fee = ltData[3]
        all_profits = ltData[4]
        self_profits = ltData[5]
        company_cd = ltData[6]
        daily_profits = ltData[7]
        #登记数据表
        result = insertAgentSum(db, stlmDate, agent_cd, txn_count, trans_amt,
                   trans_fee, all_profits, self_profits,
                   company_cd, daily_profits)
        if result != True:
            print('insertAgentSum error')
            cursor.close()
            return False
    cursor.close()
    return True

def UpdateAgentSumUpAgentDate(db, stlmDate, agent_cd, self_profts, stlm_profits):
    if len(agent_cd) < 6:
        return True
    cursor = db.cursor()
    sql = "update tbl_agent_share_sum set SUP_PROFITS = nvl(SUP_PROFITS, 0) + %f, " \
          " SUP_DAILY_PROFITS = nvl(SUP_DAILY_PROFITS, 0) + %f where host_date = '%s' and agent_cd = '%s'" \
          % (toNumberFmt(self_profts), toNumberFmt(stlm_profits), stlmDate, agent_cd)
    cursor.execute(sql)
    cursor.close()
    return True

def calcUpAgentProfits(db, stlmDate):
    sql = "select sum(self_profits),sum(STLM_PROFITS), STLM_FLAG,substr(agent_list,1,13) from tbl_agent_share_dtl where " \
          "host_date = '%s' group by STLM_FLAG, substr(agent_list,1,13) " % stlmDate
    cursor = db.cursor()
    cursor.execute(sql)
    for ltData in cursor:
        agent_list = ltData[3].split(",")
        if len(agent_list) > 1:
            self_profits = ltData[0]
            stlm_profits = ltData[1]
            stlm_flag = ltData[2]
            if stlm_flag != '2':
                stlm_profits = 0
            result = UpdateAgentSumUpAgentDate(db, stlmDate, agent_list[1], self_profits, stlm_profits)
            if result != True:
                cursor.close()
                print('UpdateAgentSumUpAgentDate error')
                return False

    cursor.close()
    return True

def addAgentAmt(agentResult, agent_cd, txn_amt):
    try:
        agentResult[agent_cd] = agentResult[agent_cd] + txn_amt
    except KeyError:
        agentResult[agent_cd] = txn_amt
    return agentResult

def getAgentAmt(agentResult, agent_cd):
    try:
        amt = agentResult[agent_cd]
    except KeyError:
        amt = 0
    return amt

def updAllAgent(db, key_rsp, allAgent, txnAgent):
    cursor = db.cursor()
    sql = "select trim(TRANS_DESCRPT) from tbl_acq_txn_log where key_rsp = '%s'" % key_rsp
    cursor.execute(sql)
    x = cursor.fetchone()
    if x is not None:
        agentList = x[0].split(",")
        i = 0
        while i <= len(agentList):
            for key in agentList[i:]:
                if len(key) < 6:
                    break
                allAgent = addAgentAmt(allAgent, agentList[i], getAgentAmt(txnAgent,key))
            i = i + 1
    cursor.close()
    return allAgent

#计算下级入账总金额
def calcLowAgentChargeProfits(db, stlmDate):
    cursor = db.cursor()
    sql = "select key_rsp, agent_cd, STLM_PROFITS from tbl_agent_share_dtl where host_date ='%s' and STLM_FLAG ='2' order by  key_rsp " % stlmDate
    cursor.execute(sql)
    lastKeyRsp = ''
    allAgent = {}
    txnAgent = {}
    for ltData in cursor:
        key_rsp = ltData[0]
        if lastKeyRsp == '':
            lastKeyRsp = key_rsp

        if lastKeyRsp != key_rsp:
            #登记总集合
            allAgent = updAllAgent(db, lastKeyRsp, allAgent, txnAgent)
            lastKeyRsp = key_rsp
            txnAgent = {}
        agent_cd = ltData[1]
        stlm_profits = ltData[2]
        txnAgent = addAgentAmt(txnAgent, agent_cd, stlm_profits)

    if lastKeyRsp != "":
        allAgent = updAllAgent(db, lastKeyRsp, allAgent, txnAgent)

    #更新数据库
    sql = "update tbl_agent_share_sum set ALL_DAILY_PROFITS = :1 where agent_cd = :2 and host_date = :3"
    cursor.prepare(sql)
    for agent_cd in allAgent:
        param = (toNumberFmt(allAgent[agent_cd]), agent_cd, stlmDate)
        cursor.execute(None, param)

    cursor.close()
    return True

def main():
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),
                           encoding='gb18030')
    if len(sys.argv) == 1:
        cursor = db.cursor()
        sql = "select BF_STLM_DATE from TBL_BAT_CUT_CTL"
        cursor.execute(sql)
        x = cursor.fetchone()
        stlm_date = x[0]
        cursor.close()
    else:
        stlm_date = sys.argv[1]

    print('hostDate %s updObjMchtSumAmt begin' % stlm_date)
    #代理商分润汇总
    if calcAgentProfitsSum(db, stlm_date) != True:
        print("calcAgentProfitsSum error")
        return False

    #计算上级利润
    if calcUpAgentProfits(db, stlm_date) != True:
        print("calcUpAgentProfits error")
        return False

    #计算下级入账总金额
    if calcLowAgentChargeProfits(db, stlm_date) != True:
        print("calcLowAgentChargeProfits error")
        return False

    db.commit()
    print('hostDate %s updObjMchtSumAmt end' % stlm_date)


if __name__ == '__main__':
    main()