#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#统计商户终端首笔

import cx_Oracle
import sys
import os
from utl.common import *

def insertData(db, host_date, company_cd, mcht_cd, term_id, term_service_flg,
                trans_amt, trans_fee, add_fee, term_ssn, key_rsp , retrivl_ref):
    cursor = db.cursor()
    #查询是否已经存在记录
    sql = "select count(*) from tbl_sn_first_txn_log where mcht_cd ='%s' and TERM_ID ='%s'" % (mcht_cd, term_id)
    cursor.execute(sql)
    x = cursor.fetchone()
    if x[0] > 0:
        return

    #查询机具号
    sql = "select sn from tbl_term_sn_key where mcht_cd ='%s' and TERM_ID ='%s'" % (mcht_cd, term_id)
    cursor.execute(sql)
    x = cursor.fetchone()
    if x is not None:
        sn = x[0]
    else:
        sn =''

    sql = "insert into tbl_sn_first_txn_log (HOST_DATE,INST_DATE,INST_TIME,SN,COMPANY_CD,MCHT_CD," \
          "TERM_ID,TERM_SERVICE_FLG,TRANS_AMT,TRANS_FEE,ADD_FEE,KEY_RSP,TERM_SSN,RETRIVL_REF) values (" \
          ":1, :2, :3, :4, :5, :6, :7, :8, :9, :10, :11, :12, :13, :14)"
    cursor.prepare(sql)
    date = getDayTime()
    param = (host_date, date[0:8], date[8:14], sn, company_cd, mcht_cd, term_id,
             term_service_flg, trans_amt, trans_fee, add_fee, key_rsp, term_ssn, retrivl_ref)
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

    print('hostDate %s recTermFirstTxn begin' % stlm_date)

    sql = "select host_date, company_cd, card_accp_id, card_accp_term_id, " \
          "substrb(PRECARD_DATD, 68, 1), nvl(trans_amt,0)/100, nvl(trans_fee,0)/100, " \
          "FEE_D_OUT, term_ssn, key_rsp , RETRIVL_REF from tbl_acq_txn_log where host_date >='%s' and " \
          "host_date <= '%s' and txn_num ='1011' and trans_state ='1' order by inst_date, inst_time " % (stlm_date, stlm_date)
    cursor = db.cursor()
    cursor.execute(sql)
    for ltData in cursor:
        host_date = ltData[0]
        company_cd = ltData[1]
        mcht_cd = ltData[2]
        term_id = ltData[3]
        if ltData[4] == '1':
            term_service_flg = 'Y'
        else:
            term_service_flg = 'N'
        trans_amt = ltData[5]
        trans_fee = ltData[6]
        add_fee = ltData[7]
        term_ssn = ltData[8]
        key_rsp = ltData[9]
        retrivl_ref = ltData[10]
        insertData(db, host_date, company_cd, mcht_cd, term_id, term_service_flg,
                   trans_amt, trans_fee, add_fee, term_ssn, key_rsp , retrivl_ref)

    cursor.close()
    db.commit()
    print('hostDate %s recTermFirstTxn end' % stlm_date)

if __name__ == '__main__':
    main()