#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#机构成本报表

import cx_Oracle
import sys
import os
from math import fabs
from openpyxl.workbook import Workbook
from utl.common import *
from utl.gldict import *

def getMchtName(mchtCd, db):
    sql = "select MCHT_NAME from TBL_OBJ_MCHT_INF where mcht_cd ='%s'" % mchtCd
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    if x is None:
        return "未知商户名"
    else:
        return x[0]

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


def newSandCostFileHead(ws):
    i = 1
    ws.cell(row=i, column=1).value = '交易日期'
    ws.cell(row=i, column=2).value = '商户号'
    ws.cell(row=i, column=3).value = '商户名'
    ws.cell(row=i, column=4).value = '结算编号'
    ws.cell(row=i, column=5).value = '项目标识'
    ws.cell(row=i, column=6).value = '商户类型'
    ws.cell(row=i, column=7).value = '借/贷'
    ws.cell(row=i, column=8).value = '交易笔数'
    ws.cell(row=i, column=9).value = '交易金额'
    ws.cell(row=i, column=10).value = '商户手续费'
    ws.cell(row=i, column=11).value = '发卡行服务费'
    ws.cell(row=i, column=12).value = '银联网络转接费'
    ws.cell(row=i, column=13).value = '品牌服务费'
    ws.cell(row=i, column=14).value = '总成本'
    ws.cell(row=i, column=15).value = '差错处理费用'
    ws.cell(row=i, column=16).value = '入账金额'
    ws.cell(row=i, column=17).value = '商户清算资金'
    ws.cell(row=i, column=18).value = '收单收入'
    ws.cell(row=i, column=19).value = '总部收入'
    ws.cell(row=i, column=20).value = '分公司收入'
    ws.cell(row=i, column=21).value = '代理收入'

def tailSandCostBody(ws,i,ltData, db):
    ws.cell(row=i, column=1).value = ltData[1]
    ws.cell(row=i, column=2).value = ltData[2]
    ws.cell(row=i, column=3).value = getMchtName(ltData[2], db)
    ws.cell(row=i, column=4).value = getStlmMd(ltData[3])
    ws.cell(row=i, column=5).value = getItemName(ltData[4])
    ws.cell(row=i, column=6).value = getMchtType(ltData[5])
    ws.cell(row=i, column=7).value = getCardType(ltData[6])
    ws.cell(row=i, column=8).value = ltData[7]
    ws.cell(row=i, column=9).value = toNumberFmt(ltData[8])
    ws.cell(row=i, column=10).value = toNumberFmt(ltData[9])
    ws.cell(row=i, column=11).value = toNumberFmt(ltData[10])
    ws.cell(row=i, column=12).value = toNumberFmt(ltData[11])
    ws.cell(row=i, column=13).value = toNumberFmt(ltData[12])
    ws.cell(row=i, column=14).value = toNumberFmt(ltData[13])
    ws.cell(row=i, column=15).value = toNumberFmt(ltData[14])
    ws.cell(row=i, column=16).value = toNumberFmt(ltData[15])
    ws.cell(row=i, column=17).value = toNumberFmt(ltData[16])
    ws.cell(row=i, column=18).value = toNumberFmt(ltData[17])
    ws.cell(row=i, column=19).value = toNumberFmt(ltData[18])
    ws.cell(row=i, column=20).value = toNumberFmt(ltData[19])
    ws.cell(row=i, column=21).value = toNumberFmt(ltData[20])

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
    if (len(stlm_date) == 8) :
        sql = "select trim(INS_ID_CD),HOST_DATE,MCHT_CD,STLM_MD,ITERM_ID, " \
              "MCHT_TYPE,CARD_TYPE,TXN_SUM,TXN_AMT,MCHT_FEE,ISS_FEE,SWT_FEE,PROD_FEE," \
              "ALL_COST_FEE,ERR_FEE,ACC_IN_AMT,MCHT_STLM_AMT,ACQ_AMT,HEAD_AMT,BRA_AMT,AGENT_AMT " \
              "from TBL_SAND_COST_FILE_DTL where host_date ='%s' order by INS_ID_CD" % stlm_date
    else:
        sql = "select trim(INS_ID_CD),HOST_DATE,MCHT_CD,STLM_MD,ITERM_ID, " \
              "MCHT_TYPE,CARD_TYPE,TXN_SUM,TXN_AMT,MCHT_FEE,ISS_FEE,SWT_FEE,PROD_FEE," \
              "ALL_COST_FEE,ERR_FEE,ACC_IN_AMT,MCHT_STLM_AMT,ACQ_AMT,HEAD_AMT,BRA_AMT,AGENT_AMT " \
              "from TBL_SAND_COST_FILE_DTL where host_date like '%s%%' order by INS_ID_CD, host_date" % stlm_date
    print(sql)
    cursor = db.cursor()
    cursor.execute(sql)
    insIdCdTmp = ''
    for ltData in cursor:
        if insIdCdTmp == '':
            insIdCdTmp = ltData[0]
            filename = filePath + 'Sand_Cost_%s_%s.xlsx' % (ltData[0], stlm_date)
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
            filename = filePath + 'Sand_Cost_%s_%s.xlsx' % (ltData[0], stlm_date)
            wb = Workbook()
            ws = wb.active
            newSandCostFileHead(ws)
            i = 2
        #写入文件
        tailSandCostBody(ws,i,ltData, db)
        i = i + 1
    if insIdCdTmp != '':
        wb.save(filename)
        wb.close()
    cursor.close()
    print('hostDate %s expRptSandCost end' % stlm_date)


if __name__ == '__main__':
    main()