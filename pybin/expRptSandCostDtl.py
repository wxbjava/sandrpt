#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#机构成本明细表

import cx_Oracle
import sys
import os
from openpyxl.workbook import Workbook
from utl.common import *


def getStlmMd(stlmMd):
    if stlmMd == '0':
        return "S0"
    elif stlmMd == '1':
        return "T1"
    else:
        return "未知"

def getMchtType(mchtType):
    if mchtType == '1':
        return "标准"
    elif mchtType == '2':
        return "优惠"
    else:
        return "其他"

def getCardType(cardType):
    if cardType == '00':
        return "贷记"
    elif cardType == '01':
        return "借记"
    else:
        return "未知"

def getAgentIncome(db, keyRsp, stlmDate):
    sql = "select PROFITS_AMT from TBL_INS_PROFITS_TXN_DTL where key_rsp = '%s' and host_date ='%s'" % (keyRsp, stlmDate)
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    if x is not None:
        return toNumberFmt(x[0])
    else:
        return 0


def newSandCostFileHead(ws):
    i = 1
    ws.cell(row=i, column=1).value = '交易日期'
    ws.cell(row=i, column=2).value = '交易时间'
    ws.cell(row=i, column=3).value = '商户号'
    ws.cell(row=i, column=4).value = '结算编号'
    ws.cell(row=i, column=5).value = '商户类型'
    ws.cell(row=i, column=6).value = '借/贷'
    ws.cell(row=i, column=7).value = '交易金额'
    ws.cell(row=i, column=8).value = '商户手续费'
    ws.cell(row=i, column=9).value = '发卡行服务费'
    ws.cell(row=i, column=10).value = '银联网络转接费'
    ws.cell(row=i, column=11).value = '品牌服务费'
    ws.cell(row=i, column=12).value = '商户清算资金'
    ws.cell(row=i, column=13).value = '收单收入'
    ws.cell(row=i, column=14).value = '总部收入'
    ws.cell(row=i, column=15).value = '代理收入'

def tailSandCostBody(ws,i,ltData, db, stlmDate):
    ws.cell(row=i, column=1).value = ltData[1]
    ws.cell(row=i, column=2).value = ltData[2]
    ws.cell(row=i, column=3).value = ltData[3]
    ws.cell(row=i, column=4).value = getStlmMd(ltData[4])
    ws.cell(row=i, column=5).value = getMchtType(ltData[5])
    ws.cell(row=i, column=6).value = getCardType(ltData[6])
    ws.cell(row=i, column=7).value = toNumberFmt(ltData[7])
    ws.cell(row=i, column=8).value = toNumberFmt(ltData[8])
    ws.cell(row=i, column=9).value = toNumberFmt(ltData[9])
    ws.cell(row=i, column=10).value = toNumberFmt(ltData[10])
    ws.cell(row=i, column=11).value = toNumberFmt(ltData[11])
    ws.cell(row=i, column=12).value = toNumberFmt(ltData[7] - ltData[8])
    ws.cell(row=i, column=13).value = toNumberFmt(ltData[8] - ltData[9] - ltData[10] - ltData[11])
    agentIncome = getAgentIncome(db, ltData[12], stlmDate)
    ws.cell(row=i, column=14).value = toNumberFmt(ltData[8] - ltData[9] - ltData[10] - ltData[11] - agentIncome)
    ws.cell(row=i, column=15).value = agentIncome

def main():
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),encoding='gb18030')
    if len(sys.argv) == 1:
        cursor = db.cursor()
        sql = "select BF_STLM_DATE from TBL_BAT_CUT_CTL"
        cursor.execute(sql)
        x = cursor.fetchone()
        stlm_date = x[0]
        cursor.close()
    else:
        stlm_date = sys.argv[1]

    print('hostDate %s expRptSandCost begin' % stlm_date)
    filePath = '%s/%s/' % (os.environ['RPT7HOME'], stlm_date)
    sql = "select trim(INS_ID_CD), txn_date, txn_time, MCHT_CD, STLM_MD, MCHT_TYPE, CARD_TYPE_2, " \
          "REAL_TRANS_AMT, mcht_fee, iss_fee, swt_fee, prod_fee, " \
          "KEY_RSP from tbl_stlm_txn_bill_dtl where host_date ='%s' and CHNL_ID='A001' order by INS_ID_CD,txn_date, txn_time" % stlm_date
    print(sql)
    cursor = db.cursor()
    cursor.execute(sql)
    insIdCdTmp = ''
    for ltData in cursor:
        if insIdCdTmp == '':
            insIdCdTmp = ltData[0]
            filename = filePath + 'Sand_Cost_Dtl_%s_%s.xlsx' % (ltData[0], stlm_date)
            wb = Workbook()
            ws = wb.active
            newSandCostFileHead(ws)
            i = 2
        if insIdCdTmp != ltData[0]:
            #关闭前一代理商文件
            wb.save(filename)
            wb.close()
            #新代理商设置文件
            insIdCdTmp = ltData[0]
            filename = filePath + 'Sand_Cost_Dtl_%s_%s.xlsx' % (ltData[0], stlm_date)
            wb = Workbook()
            ws = wb.active
            newSandCostFileHead(ws)
            i = 2
        #写入文件
        tailSandCostBody(ws,i,ltData, db, stlm_date)
        i = i + 1
    if insIdCdTmp != '':
        wb.save(filename)
        wb.close()
    cursor.close()
    print('hostDate %s expRptSandCostDtl end' % stlm_date)


if __name__ == '__main__':
    main()