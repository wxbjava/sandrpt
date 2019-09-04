#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#消费类差错交易插入差错处理表

import cx_Oracle
import sys
import os
from utl.common import *

class dbpcttxnrefundlog:
    def __init__(self):
        pass

    def setData(self, host_date, txn_num, txn_name, ori_txn_sta,
                ins_id_cd, ori_host_date, ori_key_rsp, txn_ssn,
                sand_mcht_cd, mcht_cd, pan, trans_amt, stlm_amt,
                ori_ins_profits, ori_sand_profits, ins_err_fee, txn_sta,
                credit_sta, recharge_dt):
        self.host_date = host_date
        self.txn_num = txn_num
        self.txn_name = txn_name
        self.ori_txn_sta = ori_txn_sta
        self.ins_id_cd = ins_id_cd
        self.ori_host_date = ori_host_date
        self.ori_key_rsp = ori_key_rsp
        self.txn_ssn = txn_ssn
        self.sand_mcht_cd = sand_mcht_cd
        self.mcht_cd = mcht_cd
        self.pan = pan
        self.trans_amt = trans_amt
        self.stlm_amt = stlm_amt
        self.ori_ins_profits = ori_ins_profits
        self.ori_sand_profits = ori_sand_profits
        self.ins_err_fee = ins_err_fee
        self.txn_sta = txn_sta
        self.credit_sta = credit_sta
        self.recharge_dt = recharge_dt

    def insertdb(self, db):
        date = getDayTime()
        sql = "INSERT INTO TBL_PCT_TXN_REFUND_LOG (INST_DATE, INST_TIME, HOST_DATE, TXN_NUM, TXN_NAME, " \
              "ORI_TXN_STA, INS_ID_CD, ORI_HOST_DATE, ORI_KEY_RSP, TXN_SSN, SAND_MCHT_CD, MCHT_CD, PAN, " \
              "TRANS_AMT, STLM_AMT, ORI_INS_PROFITS, ORI_SAND_PROFITS, INS_ERR_FEE, TXN_STA, CREDIT_STA, " \
              "RECHARGE_DT, REC_OPR_ID, REC_UPD_OPR, REC_CRT_TS, REC_UPD_TS) VALUES " \
              "('%s', '%s', :1, :2, :3, :4, :5, :6, " \
              ":7, :8," \
              " :9, :10, :11, " \
              ":12, :13, :14, :15, :16, :17, :18, :19, null, null, sysdate, sysdate)" % (date[0:8], date[8:14])
        cursor = db.cursor()
        cursor.prepare(sql)
        param = (self.host_date,self.txn_num,self.txn_name,self.ori_txn_sta,self.ins_id_cd,
                 self.ori_host_date,self.ori_key_rsp,self.txn_ssn,self.sand_mcht_cd,
                 self.mcht_cd,self.pan,self.trans_amt,self.stlm_amt,self.ori_ins_profits,
                 self.ori_sand_profits,self.ins_err_fee,self.txn_sta,self.credit_sta,self.recharge_dt)
        cursor.execute(None, param)
        cursor.close()
        return True


def handrtntxn(db, txnData):
    #扣款类型交易
    cursor = db.cursor()
    #查找原交易
    sql = "select key_rsp, ISS_FEE, SWT_FEE, PROD_FEE, INS_ID_CD,mcht_cd,pan from TBL_STLM_TXN_BILL_DTL where txn_key ='%s' and CHNL_ID ='A001'" % txnData[0][:12]
    cursor.execute(sql)
    x = cursor.fetchone()
    if x is None:
        print('get ori stlm_txn_bill error')
        return

    ins_err_fee = toNumberFmt(x[1] + x[2] + x[3] + txnData[3] + txnData[4] + txnData[5] + txnData[6]) #一般为品牌服务费+差错费

    #查找原分润
    sql = "select mcht_cd, TRANS_AMT, TRANS_FEE, PROFITS_AMT,host_date from TBL_INS_PROFITS_TXN_DTL where key_rsp = '%s'" % x[0]
    cursor.execute(sql)
    y = cursor.fetchone()
    if y is None:
        print('get ori ins_profits error')
        return
    ori_sand_profits = toNumberFmt(y[2] - (x[1] + x[2] + x[3]) - y[3])
    stlm_amt = toNumberFmt(y[1] - y[2])

    dblog = dbpcttxnrefundlog()
    dblog.setData(host_date=getDayTime()[0:8],
                  txn_num=txnData[1],
                  txn_name=txnData[2],
                  ori_txn_sta='1',
                  ins_id_cd=x[4],
                  ori_host_date=y[4],
                  ori_key_rsp=x[0],
                  txn_ssn=txnData[0],
                  sand_mcht_cd=x[5],
                  mcht_cd=y[0],
                  pan=x[6],
                  trans_amt=y[1],
                  stlm_amt=stlm_amt,
                  ori_ins_profits=y[3],
                  ori_sand_profits=ori_sand_profits,
                  ins_err_fee=ins_err_fee,
                  txn_sta='0',
                  credit_sta='0',
                  recharge_dt=''
                  )
    dblog.insertdb(db)

def main():
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),
                           encoding='gb18030')
    stlm_date = sys.argv[1]
    #9005-发卡退单
    #9008-请款
    #9009-手工退货
    sql = "select txn_key,TXN_NUM, TXN_DSP, ISS_FEE, SWT_FEE, PROD_FEE, ERR_FEE from TBL_STLM_TXN_BILL_DTL where stlm_date ='%s' and CHNL_ID ='A001' " \
          " and txn_num in ('9005','9008','9009')" % stlm_date
    cursor = db.cursor()
    cursor.execute(sql)
    for ltData in cursor:
        if ltData[1] == '9005' or ltData[1] == '9009':
            handrtntxn(db, ltData)

    cursor.close()
    db.commit()
    db.close()

if __name__ == '__main__':
    main()