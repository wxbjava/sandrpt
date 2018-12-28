#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#代付勾兑完成后,更新代付通道信息

import cx_Oracle
import os
import sys

def updObjChnlId(db, destChnlId, keyRsp):
    sql = "update tbl_acq_txn_log set RCV_INS_ID_CD = '%s' where key_rsp = '%s'" % (destChnlId, keyRsp)
    upd = db.cursor()
    upd.execute(sql)
    db.commit()
    upd.close()

def main():
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),
                           encoding='gb18030')
    # 获取清算日
    if len(sys.argv) == 1:
        cursor = db.cursor()
        sql = "select BF_STLM_DATE from TBL_BAT_CUT_CTL"
        cursor.execute(sql)
        x = cursor.fetchone()
        stlm_date = x[0]
        cursor.close()
    else:
        stlm_date = sys.argv[1]

    print('hostDate %s update paytxn objinfo begin' % stlm_date)

    sql = "select DEST_CHNL_ID, KEY_RSP from TBL_STLM_TXN_BILL_DTL " \
          "where CHNL_ID ='A002' and txn_num ='1801' and CHECK_STA ='1' and host_date ='%s'" % stlm_date
    cursor = db.cursor()
    cursor.execute(sql)
    for ltTxn in cursor:
        destChnlId = ltTxn[0]
        keyRsp = ltTxn[1]
        updObjChnlId(db, destChnlId, keyRsp)
    cursor.close()
    db.close()

    print('hostDate %s update paytxn objinfo end' % stlm_date)

if __name__ == '__main__':
    main()