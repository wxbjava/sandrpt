#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#每日日切后统计银联商户累计金额

import cx_Oracle
import os
import sys
from utl.common import *

def updSandMchtAmtSum(db, mcht_cd, amt):
    sql = "update tbl_obj_mcht_inf set SUM_AMT = nvl(SUM_AMT,0) + %.f  where mcht_cd ='%s'" % (toNumberFmt(amt), mcht_cd)
    cursor = db.cursor()
    cursor.execute(sql)
    if cursor.rowcount == 0:
        #获取商户名
        sql = "select trim(MCHT_NM) from tbl_shc_mcht_pool where mcht_cd = '%s'" % mcht_cd
        cursor.execute(sql)
        x = cursor.fetchone()
        if x is None:
            mcht_nm = '未知商户名'
        else:
            mcht_nm = x[0]
        #插入
        sql = "insert into tbl_obj_mcht_inf  (MCHT_CD, MCHT_NAME, SUM_AMT)  values " \
              "(:1, :2, :3)"
        cursor.prepare(sql)
        param = (mcht_cd, mcht_nm, toNumberFmt(amt))
        cursor.execute(None, param)
    cursor.close()


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
    sql = "select sum(trans_amt/100), obj_mcht_cd from tbl_acq_txn_log where host_date ='%s' " \
          " and txn_num ='1011' and trans_state = '1' group by obj_mcht_cd" % stlm_date
    cursor = db.cursor()
    cursor.execute(sql)
    for ltData in cursor:
        amt = ltData[0]
        mcht_cd = ltData[1]
        #登记数据库表
        updSandMchtAmtSum(db, mcht_cd, amt)


    print('hostDate %s updObjMchtSumAmt end' % stlm_date)
    cursor.close()
    db.commit()
    db.close()

if __name__ == '__main__':
    main()