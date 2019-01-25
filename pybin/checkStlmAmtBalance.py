#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#����ƽ���,�˶������ʽ��������ʽ��ϵ

import os
import sys
import cx_Oracle
from openpyxl.workbook import Workbook
from utl.common import *


#���㱾������������
def calcStlmInnerAmt(db, stlm_date):
    sql = "select sum(trans_amt)/100 from tbl_acq_txn_log where host_date = '%s' and txn_num ='1011' " \
          "and trans_state ='1' and REVSAL_FLAG ='0' and CANCEL_FLAG ='0'" % stlm_date
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    if x is not None:
        return toNumberFmt(x[0])
    else:
        return 0

#�������ճ����ʽ�
def calcLastChnlStlmFunds(db, stlm_date):
    sql = "select sum(REAL_TRANS_AMT) from tbl_stlm_txn_bill_dtl where host_date ='%s' and " \
          "stlm_date = '%s' and check_sta ='1' and txn_num ='1011'"\
          % (stlm_date, getLastDay(stlm_date))
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    if x is not None:
        return toNumberFmt(x[0])
    else:
        return 0

#���㱾�ճ����ʽ�
def calcChnlStlmFunds(db, stlm_date):
    #ͨ���ļ�����,��˾�����մ���
    sql = "select sum(REAL_TRANS_AMT) from tbl_stlm_txn_bill_dtl where host_date ='%s' and " \
          "stlm_date = '%s' and check_sta ='1' and txn_num ='1011'" \
          % (getNextDay(stlm_date), stlm_date)
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    if x is not None:
        return toNumberFmt(x[0])
    else:
        return 0

#���㵱�ղ���ʽ�
def calcErrAmt(db, stlm_date):
    sql = "select sum(REAL_TRANS_AMT) from tbl_stlm_txn_bill_dtl where stlm_date ='%s' and " \
          "txn_num !='1011' and chnl_id ='A001'" % stlm_date
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    if x is not None:
        return toNumberFmt(x[0])
    else:
        return 0

#������˳���׽��
def calcLongTxnAmt(db, stlm_date):
    sql = "select sum(CHNL_TXN_AMT) from tbl_err_chk_txn_dtl where " \
          "host_date = '%s' and CHK_STA ='1' and group_id ='A001'" % stlm_date
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    if x is not None:
        return toNumberFmt(x[0])
    else:
        return 0

#����ͨ�������ļ��ܽ��
def calcChnlAmt(db, stlm_date):
    sql = "select sum(REAL_TRANS_AMT) from TBL_STLM_TXN_BILL_DTL where stlm_date ='%s' and chnl_id ='A001'" % stlm_date
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    if x is not None:
        return toNumberFmt(x[0])
    else:
        return 0


def main():
    # ���ݿ���������
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),encoding='gb18030')
    # ��ȡ������
    if len(sys.argv) == 1:
        cursor = db.cursor()
        sql = "select BF_STLM_DATE from TBL_BAT_CUT_CTL"
        cursor.execute(sql)
        x = cursor.fetchone()
        stlm_date = x[0]
        cursor.close()
    else:
        stlm_date = sys.argv[1]
    print('hostDate %s rpt begin' % stlm_date)

    innerAmt = calcStlmInnerAmt(db, stlm_date)
    lastChnlFunds = calcLastChnlStlmFunds(db, stlm_date)
    chnlFunds = calcChnlStlmFunds(db, stlm_date)
    longAmt = calcLongTxnAmt(db, stlm_date)
    chnlAmt = calcChnlAmt(db, stlm_date)
    errAmt = calcErrAmt(db, stlm_date)

    print("innerAmt:%.2f" % innerAmt)
    print("lastChnlFunds:%.2f" % lastChnlFunds)
    print("chnlFunds:%.2f" % chnlFunds)
    print("longAmt:%.2f" % longAmt)
    print("errAmt:%.2f" % errAmt)
    print("chnlAmt:%.2f" % chnlAmt)

    if innerAmt - lastChnlFunds + chnlFunds + errAmt + longAmt != chnlAmt:
        bal_sta = '2'
    else:
        bal_sta = '1'

    sql = "insert into TBL_STLM_TASK_CTL (host_date, " \
          "chnl_amt, " \
          "bal_mark) values ('%s', %.2f, '%s')" % (stlm_date, chnlAmt, bal_sta)
    cursor = db.cursor()
    cursor.execute(sql)
    db.commit()
    cursor.close()

    #���������Ҫ�����ļ�
    wb = Workbook()
    ws = wb.active

    ws.cell(row=1, column=8).value = '�յ�����ϵͳ�����ʽ�ƽ���'
    i = 2
    ws.cell(row=i, column=1).value = '���������ܽ��'
    ws.cell(row=i, column=2).value = '���ճ����ʽ�'
    ws.cell(row=i, column=3).value = '���ճ����ʽ�'
    ws.cell(row=i, column=4).value = '����ཻ�׽��'
    ws.cell(row=i, column=5).value = '�����׽��'
    ws.cell(row=i, column=6).value = 'ͨ���ļ����'
    ws.cell(row=i, column=7).value = '�˶Խ��'
    i = i + 1
    ws.cell(row=i, column=1).value = innerAmt
    ws.cell(row=i, column=2).value = lastChnlFunds
    ws.cell(row=i, column=3).value = chnlFunds
    ws.cell(row=i, column=4).value = errAmt
    ws.cell(row=i, column=5).value = longAmt
    ws.cell(row=i, column=6).value = chnlAmt
    if bal_sta == '1':
        ws.cell(row=i, column=7).value = 'ƽ��'
    else:
        ws.cell(row=i, column=7).value = '��ƽ'

    filePath = '%s/%s/' % (os.environ['RPT7HOME'], stlm_date)
    filename = filePath + 'AcqStlmCheckFile01_%s.xlsx' % stlm_date
    wb.save(filename)

    db.close()

if __name__ == '__main__':
    main()