#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#机构成本报表

import cx_Oracle
import sys
import os
import xlsxwriter
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
    elif mchtType == '5':
        return "云闪付"
    else:
        return "其他"

def getCardType(cardType):
    if cardType == '00':
        return "贷记"
    elif cardType == '01':
        return "借记"
    else:
        return "未知"


def newSandCostFileHead(ws, i):
    data = []
    data.append('交易日期')
    data.append('商户号')
    data.append('商户名')
    data.append('结算编号')
    data.append('项目标识')
    data.append('商户类型')
    data.append('借/贷')
    data.append('交易笔数')
    data.append('交易金额')
    data.append('商户手续费')
    data.append('发卡行服务费')
    data.append('银联网络转接费')
    data.append('品牌服务费')
    data.append('总成本')
    data.append('差错处理费用')
    data.append('入账金额')
    data.append('商户清算资金')
    data.append('收单收入')
    data.append('总部收入')
    data.append('分公司收入')
    data.append('代理收入')
    ws.write_row(i, 0, data)

def tailSandCostBody(ws, ltData, db, i):
    data= []
    data.append(ltData[1])
    data.append(ltData[2])
    data.append(getMchtName(ltData[2], db))
    data.append(getStlmMd(ltData[3]))
    data.append(getItemName(ltData[4]))
    data.append(getMchtType(ltData[5]))
    data.append(getCardType(ltData[6]))
    data.append(ltData[7])
    data.append(toNumberFmt(ltData[8]))
    data.append(toNumberFmt(ltData[9]))
    data.append(toNumberFmt(ltData[10]))
    data.append(toNumberFmt(ltData[11]))
    data.append(toNumberFmt(ltData[12]))
    data.append(toNumberFmt(ltData[13]))
    data.append(toNumberFmt(ltData[14]))
    data.append(toNumberFmt(ltData[15]))
    data.append(toNumberFmt(ltData[16]))
    data.append(toNumberFmt(ltData[17]))
    data.append(toNumberFmt(ltData[18]))
    data.append(toNumberFmt(ltData[19]))
    data.append(toNumberFmt(ltData[20]))
    ws.write_row(i, 0, data)

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
    i = 0
    for ltData in cursor:
        if insIdCdTmp == '':
            insIdCdTmp = ltData[0]
            filename = filePath + 'Sand_Cost_%s_%s.xlsx' % (ltData[0], stlm_date)
            wb = xlsxwriter.Workbook(filename, {'constant_memory': True})
            ws = wb.add_worksheet('机构成本')
            newSandCostFileHead(ws, i)
            i = i + 1
        if insIdCdTmp != ltData[0]:
            #关闭前一代理商文件
            wb.close()
            i = 0
            #新代理商设置文件
            insIdCdTmp = ltData[0]
            filename = filePath + 'Sand_Cost_%s_%s.xlsx' % (ltData[0], stlm_date)
            wb = xlsxwriter.Workbook(filename, {'constant_memory': True})
            ws = wb.add_worksheet('机构成本')
            newSandCostFileHead(ws, i)
            i = i + 1
        #写入文件
        tailSandCostBody(ws,ltData, db, i)
        i = i + 1
    if insIdCdTmp != '':
        wb.close()
    cursor.close()
    print('hostDate %s expRptSandCost end' % stlm_date)


if __name__ == '__main__':
    main()