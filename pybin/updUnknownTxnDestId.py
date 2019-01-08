#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#日切后初始操作,将核心的未知交易从前置更新通道信息，以便代付对账


import cx_Oracle
import os
import sys


def updDestId(db, keyRsp):
    sql = "select dest_bin, trim(obj_mcht_cd), trim(OBJ_TERM_CD) from tbl_shc_log1 where key_rsp = '%s'" % keyRsp
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    if x is not None:
        sql = "update tbl_acq_txn_log set obj_ins_id_cd = '%s',OBJ_MCHT_CD='%s',OBJ_TERM_CD='%s' where " \
              "key_rsp = '%s'" % (x[0],x[1],x[2],keyRsp)
    else:
        sql = "update tbl_acq_txn_log set obj_ins_id_cd = '0000' where " \
              "key_rsp = '%s'" % keyRsp
    upd = db.cursor()
    upd.execute(sql)
    db.commit()
    upd.close()
    cursor.close()


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

    print('hostDate %s update unknown destid begin' % stlm_date)

    sql = "select key_rsp from tbl_acq_txn_log where host_date ='%s' and txn_num in ('1801','1011') and trans_state ='3'" % stlm_date
    cursor = db.cursor()
    cursor.execute(sql)
    for ltTxn in cursor:
        keyRsp = ltTxn[0]
        updDestId(db, keyRsp)
    cursor.close()
    db.close()

if __name__ == '__main__':
    main()