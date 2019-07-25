#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#每日日切后统计银联商户累计金额



import cx_Oracle
import os
import sys

def updSandMchtAmtSum(db, mcht_cd, amt):
    sql = "update tbl_obj_mcht_inf set "


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



    print('hostDate %s updObjMchtSumAmt end' % stlm_date)

if __name__ == '__main__':
    main()