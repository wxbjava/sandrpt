#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#对账平衡表,核对来账资金与自主资金关系

import os
import sys
import cx_Oracle
from openpyxl.workbook import Workbook
from utl.common import *


#计算本日自主清算金额
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

#计算上日沉淀资金
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

#计算本日沉淀资金
def calcChnlStlmFunds(db, stlm_date):
    #通道文件当日,我司清算日次日
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

#计算当日差错资金
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

#计算对账长款交易金额
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

#计算通道来账文件总金额
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
    # 数据库连接配置
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),encoding='gb18030')
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

    #结果数据需要生成文件
    wb = Workbook()
    ws = wb.active

    ws.cell(row=1, column=8).value = '收单间联系统对账资金平衡表'
    i = 2
    ws.cell(row=i, column=1).value = '自主清算总金额'
    ws.cell(row=i, column=2).value = '上日沉淀资金'
    ws.cell(row=i, column=3).value = '当日沉淀资金'
    ws.cell(row=i, column=4).value = '差错类交易金额'
    ws.cell(row=i, column=5).value = '长交易金额'
    ws.cell(row=i, column=6).value = '通道文件金额'
    ws.cell(row=i, column=7).value = '核对结果'
    i = i + 1
    ws.cell(row=i, column=1).value = innerAmt
    ws.cell(row=i, column=2).value = lastChnlFunds
    ws.cell(row=i, column=3).value = chnlFunds
    ws.cell(row=i, column=4).value = errAmt
    ws.cell(row=i, column=5).value = longAmt
    ws.cell(row=i, column=6).value = chnlAmt
    if bal_sta == '1':
        ws.cell(row=i, column=7).value = '平衡'
    else:
        ws.cell(row=i, column=7).value = '不平'

    filePath = '%s/%s/' % (os.environ['RPT7HOME'], stlm_date)
    filename = filePath + 'AcqStlmCheckFile01_%s.xlsx' % stlm_date
    wb.save(filename)

    db.close()

if __name__ == '__main__':
    main()