#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#间联系统对账报表
#可传参自主交易日,生成指定日志报表

import cx_Oracle
import sys
from openpyxl.workbook import Workbook
import os
from utl.common import *
from utl.gldict import *


#表格全局量
i = 0

#通过键值查找一代的总分润
def getAgentIncome(db, key_rsp):
    cursor = db.cursor()
    sql = "select profits_amt from tbl_ins_profits_txn_dtl where key_rsp = '%s'" % key_rsp
    cursor.execute(sql)
    x = cursor.fetchone()
    if x is not None:
        amt = toNumberFmt(x[0])
    else:
        amt = 0.0
    cursor.close()
    return amt

#通过键值查找通道交易相关信息
def getChnlTxn(db, chnlId, chlTxnKey):
    cursor = db.cursor()
    sql = "select MCHT_CD, TERM_CD, REAL_TRANS_AMT, ISS_FEE, SWT_FEE, PROD_FEE  " \
          "from TBL_STLM_TXN_BILL_DTL where CHNL_ID = '%s' and TXN_KEY = '%s'" % (chnlId, chlTxnKey)
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    return x[0], x[1], toNumberFmt(x[2]), toNumberFmt(x[3]), toNumberFmt(x[4]), toNumberFmt(x[5])

#通过键值查找交易信息,返回商户号,终端号,交易金额
def getOwnTxn(db, keyRsp) :
    cursor = db.cursor()
    sql = "select card_accp_id, CARD_ACCP_TERM_ID, trans_amt  " \
          "from tbl_acq_txn_log where key_rsp = '%s'" %  keyRsp
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    return x[0], x[1], toNumberFmt(int(x[2]))

#消费类对账交易
def handleTxn01Rpt(db, ws, stlm_date):
    handleTxn01RptHead(ws)
    handleTxn01RptBody(db, ws, stlm_date)

def handleTxn01RptHead(ws):
    #报表名称
    ws.cell(row=1, column=8).value = '收单间联清算对账报表'
    #报表头
    ws.cell(row=2, column=1).value = '系统交易日期'
    ws.cell(row=2, column=2).value = '项目标识'
    ws.cell(row=2, column=3).value = '交易笔数'
    ws.cell(row=2, column=4).value = '交易金额'
    ws.cell(row=2, column=5).value = '商户手续费'
    ws.cell(row=2, column=6).value = '发卡服务费'
    ws.cell(row=2, column=7).value = '银联网络转接费'
    ws.cell(row=2, column=8).value = '品牌服务费'
    ws.cell(row=2, column=9).value = '总成本'
    ws.cell(row=2, column=10).value = '差错处理费用'
    ws.cell(row=2, column=11).value = '资金清算净额'
    ws.cell(row=2, column=12).value = '收单收入'
    ws.cell(row=2, column=13).value = '总部收入'
    ws.cell(row=2, column=14).value = '分公司收入'
    ws.cell(row=2, column=15).value = '代理商收入'


def handleTxn01RptBody(db, ws, stlm_date):
    count,allCount = 0, 0
    transAmt,allTransAmt = 0.0, 0.0
    issAmt,allIssAmt = 0.0, 0.0
    swtAmt,allSwtAmt = 0.0, 0.0
    prodAmt,allProdAmt = 0.0, 0.0
    errAmt,allErrAmt = 0.0, 0.0
    mchtFee,allMchtFee = 0.0, 0.0
    companyIncome,allCompanyIncome = 0.0, 0.0
    agentIncome,allAgentIncome = 0.0, 0.0
    old_itemId = ''
    itemId = ''
    #查找对账成功的交易
    global i
    i = 3
    sql = "select ITEM_ID, key_rsp, REAL_TRANS_AMT, mcht_fee, " \
          "iss_fee, swt_fee, prod_fee, err_fee from TBL_STLM_TXN_BILL_DTL " \
          "where host_date = '%s' and txn_num ='1011' and check_sta ='1' order by ITEM_ID " % (stlm_date)
    cursor = db.cursor()
    cursor.execute(sql)
    for ltTxn in cursor:
        if count == 0:
            old_itemId = ltTxn[0]
        itemId = ltTxn[0]
        if (old_itemId != itemId):
            #项目id不同
            tailTxn01RptBody(ws,i,stlm_date,old_itemId,count,transAmt,issAmt,swtAmt,
                             prodAmt,errAmt,mchtFee,companyIncome,agentIncome)
            i = i + 1
            #初始化
            old_itemId = itemId
            count = 0
            transAmt = 0.0
            issAmt = 0.0
            swtAmt = 0.0
            prodAmt = 0.0
            errAmt = 0.0
            mchtFee = 0.0
            companyIncome = 0.0
            agentIncome = 0.0

        #查找代理商费用
        key_rsp = ltTxn[1]
        amtTmp = getAgentIncome(db, key_rsp)
        agentIncome = toNumberFmt(agentIncome + amtTmp)
        allAgentIncome = toNumberFmt(allAgentIncome + amtTmp)
        count = count + 1
        allCount = allCount + 1
        transAmt = toNumberFmt(transAmt + ltTxn[2])
        allTransAmt = toNumberFmt(allTransAmt + ltTxn[2])
        mchtFee = toNumberFmt(mchtFee + ltTxn[3])
        allMchtFee = toNumberFmt(allMchtFee + ltTxn[3])
        issAmt = toNumberFmt(issAmt + ltTxn[4])
        allIssAmt = toNumberFmt(allIssAmt + ltTxn[4])
        swtAmt = toNumberFmt(swtAmt + ltTxn[5])
        allSwtAmt = toNumberFmt(allSwtAmt + ltTxn[5])
        prodAmt = toNumberFmt(prodAmt + ltTxn[6])
        allProdAmt = toNumberFmt(allProdAmt + ltTxn[6])
        errAmt = toNumberFmt(errAmt + ltTxn[7])
        allErrAmt = toNumberFmt(allErrAmt + ltTxn[7])
    if itemId != '':
        tailTxn01RptBody(ws, i, stlm_date, old_itemId, count, transAmt, issAmt, swtAmt,
                         prodAmt, errAmt, mchtFee, companyIncome, agentIncome)
        i = i + 1

    tailTxn01RptTail(ws, i, allCount, allTransAmt, allIssAmt, allSwtAmt,
                     allProdAmt, allErrAmt, allMchtFee, allCompanyIncome, allAgentIncome)
    i = i + 1

    cursor.close()

def tailTxn01RptBody(ws,i,stlmDate,itemId,count,transAmt,issAmt,swtAmt,
                    prodAmt,errAmt,mchtFee,companyIncome,agentIncome) :
    ws.cell(row=i, column=1).value = stlmDate
    ws.cell(row=i, column=2).value = getItemName(itemId)
    ws.cell(row=i, column=3).value = count
    ws.cell(row=i, column=4).value = transAmt
    ws.cell(row=i, column=5).value = mchtFee
    ws.cell(row=i, column=6).value = issAmt
    ws.cell(row=i, column=7).value = swtAmt
    ws.cell(row=i, column=8).value = prodAmt
    ws.cell(row=i, column=9).value = toNumberFmt(issAmt + swtAmt + prodAmt)
    ws.cell(row=i, column=10).value = errAmt
    ws.cell(row=i, column=11).value = toNumberFmt(transAmt - (issAmt + swtAmt + prodAmt))
    ws.cell(row=i, column=12).value = toNumberFmt(mchtFee - (issAmt + swtAmt + prodAmt))
    ws.cell(row=i, column=13).value = toNumberFmt(mchtFee - (issAmt + swtAmt + prodAmt) - companyIncome - agentIncome)
    ws.cell(row=i, column=14).value = companyIncome
    ws.cell(row=i, column=15).value = agentIncome

def tailTxn01RptTail(ws,i,count,transAmt,issAmt,swtAmt,
                    prodAmt,errAmt,mchtFee,companyIncome,agentIncome) :
    ws.cell(row=i, column=2).value = '小计'
    ws.cell(row=i, column=3).value = count
    ws.cell(row=i, column=4).value = transAmt
    ws.cell(row=i, column=5).value = mchtFee
    ws.cell(row=i, column=6).value = issAmt
    ws.cell(row=i, column=7).value = swtAmt
    ws.cell(row=i, column=8).value = prodAmt
    ws.cell(row=i, column=9).value = toNumberFmt(issAmt + swtAmt + prodAmt)
    ws.cell(row=i, column=10).value = errAmt
    ws.cell(row=i, column=11).value = toNumberFmt(transAmt - (issAmt + swtAmt + prodAmt))
    ws.cell(row=i, column=12).value = toNumberFmt(mchtFee - (issAmt + swtAmt + prodAmt))
    ws.cell(row=i, column=13).value = toNumberFmt(mchtFee - (issAmt + swtAmt + prodAmt) - companyIncome - agentIncome)
    ws.cell(row=i, column=14).value = companyIncome
    ws.cell(row=i, column=15).value = agentIncome


#代付类对账交易
def handleTxn02Rpt(db, ws, stlm_date):
    handleTxn02RptHead(ws)
    handleTxn02RptBody(db, ws, stlm_date)

def handleTxn02RptHead(ws):
    global i
    i = i + 2
    #报表头
    ws.cell(row=i, column=1).value = '代付日期'
    ws.cell(row=i, column=2).value = 'S0/T1'
    ws.cell(row=i, column=3).value = '代付笔数'
    ws.cell(row=i, column=4).value = '代付金额'
    ws.cell(row=i, column=5).value = '品牌服务费'
    ws.cell(row=i, column=6).value = '代付成本'
    ws.cell(row=i, column=7).value = '资金清算净额'
    i = i + 1


def handleTxn02RptBody(db, ws, stlm_date):
    count,allCount = 0, 0
    transAmt, allTransAmt= 0.0, 0.0
    prodAmt, allProdAmt = 0.0, 0.0
    costAmt, allCostAmt = 0.0, 0.0
    oldPayType = ''
    payType = ''
    #查找对账成功的交易
    global i
    sql = "select pay_type, REAL_TRANS_AMT, prod_fee, ISS_FEE " \
          " from TBL_STLM_TXN_BILL_DTL " \
          "where host_date = '%s' and txn_num ='1801' and check_sta ='1' order by  pay_type" % (stlm_date)
    cursor = db.cursor()
    cursor.execute(sql)
    for ltTxn in cursor:
        if count == 0:
            oldPayType = ltTxn[0]
        payType = ltTxn[0]
        if (oldPayType != payType):
            #代付商户不同
            tailTxn02RptBody(ws,i,stlm_date,oldPayType,count,transAmt,prodAmt,costAmt)
            i = i + 1
            #初始化
            oldPayType = payType
            count = 0
            transAmt = 0.0
            prodAmt = 0.0
            costAmt = 0.0

        #查找代理商费用
        count = count + 1
        allCount = allCount + 1
        transAmt = toNumberFmt(transAmt + ltTxn[1])
        allTransAmt = toNumberFmt(allTransAmt + ltTxn[1])
        prodAmt = toNumberFmt(prodAmt + ltTxn[2])
        allProdAmt = toNumberFmt(allProdAmt + ltTxn[2])
        costAmt = toNumberFmt(costAmt + ltTxn[3])
        allCostAmt = toNumberFmt(allCostAmt + ltTxn[3])
    if payType != '':
        tailTxn02RptBody(ws, i, stlm_date, payType, count, transAmt, prodAmt, costAmt)
        i = i + 1
    tailTxn02RptTail(ws, i, allCount, allTransAmt, allProdAmt, allCostAmt)
    i = i + 1

    cursor.close()



def tailTxn02RptBody(ws,i,stlmDate,payType,count,transAmt,prodAmt,costAmt):
    ws.cell(row=i, column=1).value = stlmDate
    if payType == '00':
        ws.cell(row=i, column=2).value = 'S0出款'
    elif payType == '01':
        ws.cell(row=i, column=2).value = 'T1出款'
    else:
        ws.cell(row=i, column=2).value = '其他'
    ws.cell(row=i, column=3).value = count
    ws.cell(row=i, column=4).value = toNumberFmt(transAmt * (-1))
    ws.cell(row=i, column=5).value = prodAmt
    ws.cell(row=i, column=6).value = costAmt
    ws.cell(row=i, column=7).value = prodAmt
    ws.cell(row=i, column=7).value = toNumberFmt((transAmt - prodAmt - costAmt) * (-1))

def tailTxn02RptTail(ws,i,count,transAmt,prodAmt,costAmt):
    ws.cell(row=i, column=2).value = '总计'
    ws.cell(row=i, column=3).value = count
    ws.cell(row=i, column=4).value = toNumberFmt(transAmt * (-1))
    ws.cell(row=i, column=5).value = prodAmt
    ws.cell(row=i, column=6).value = costAmt
    ws.cell(row=i, column=7).value = prodAmt
    ws.cell(row=i, column=7).value = toNumberFmt((transAmt - prodAmt - costAmt) * (-1))

#出款
def tailTxn03RptBody(ws,i,stlm_date,chnlId,count,transAmt):
    ws.cell(row=i, column=1).value = stlm_date
    ws.cell(row=i, column=2).value = getChnlName(chnlId)
    ws.cell(row=i, column=3).value = count
    ws.cell(row=i, column=4).value = toNumberFmt(transAmt * (-1))
    if chnlId.rstrip() == '00000901':
        ws.cell(row=i, column=5).value = '代理商分润'
    else:
        ws.cell(row=i, column=5).value = '商户清算款'

def handleTxn03RptBody(db, ws, stlm_date):
    # 按照通道查找对账成功的代付
    global i
    sql = "select DEST_CHNL_ID, count(*), sum(REAL_TRANS_AMT) " \
          " from TBL_STLM_TXN_BILL_DTL " \
          "where host_date = '%s' and txn_num ='1801' and check_sta ='1' group by DEST_CHNL_ID" % (stlm_date)
    cursor = db.cursor()
    cursor.execute(sql)
    for ltTxn in cursor:
        tailTxn03RptBody(ws, i, stlm_date, ltTxn[0], ltTxn[1], toNumberFmt(ltTxn[2]))
        i = i + 1

    cursor.close()

def handleTxn03Rpt(db, ws, stlm_date):
    global i
    i = i + 2
    ws.cell(row=i, column=1).value = '出款'
    # 报表头
    i = i + 1
    ws.cell(row=i, column=1).value = '出款日期'
    ws.cell(row=i, column=2).value = '通道'
    ws.cell(row=i, column=3).value = '出款笔数'
    ws.cell(row=i, column=4).value = '出款金额'
    ws.cell(row=i, column=5).value = '出款类型'
    i = i + 1

    handleTxn03RptBody(db, ws, stlm_date)


#POS长短款交易明细
def handleTxn04Rpt(db, ws, stlm_date):
    handleTxn04RptHead(ws)
    handleTxn04RptBody(db, ws, stlm_date)

def handleTxn04RptHead(ws):
    global i
    i = i + 2
    ws.cell(row=i, column=1).value = 'POS长短款交易明细'
    #报表头
    i = i + 1
    ws.cell(row=i, column=2).value = '交易日期'
    ws.cell(row=i, column=3).value = '交易时间'
    ws.cell(row=i, column=4).value = '类别'
    ws.cell(row=i, column=5).value = '交易商户号'
    ws.cell(row=i, column=6).value = '交易终端号'
    ws.cell(row=i, column=7).value = '商户号'
    ws.cell(row=i, column=8).value = '终端号'
    ws.cell(row=i, column=9).value = '帐号'
    ws.cell(row=i, column=10).value = '交易金额'
    ws.cell(row=i, column=11).value = '发卡服务费'
    ws.cell(row=i, column=12).value = '银联服务费'
    ws.cell(row=i, column=13).value = '品牌服务费'
    ws.cell(row=i, column=14).value = '资金清算净额'

    i = i + 1


def handleTxn04RptBody(db, ws, stlm_date):
    #查找长短款交易
    global i
    sql = "select trans_date, trans_time, CHK_STA, key_rsp, CHNL_RETRIVL_REF, GROUP_ID, " \
          " pan " \
          " from TBL_ERR_CHK_TXN_DTL " \
          "where host_date = '%s' and txn_num ='1011' and CHK_STA ='1' order by  CHK_STA" % (stlm_date)
    cursor = db.cursor()
    cursor.execute(sql)
    for ltTxn in cursor:
        transDate = ltTxn[0]
        transTime = ltTxn[1]
        chkSta = ltTxn[2]
        pan = ltTxn[6]
        if chkSta == '1':
            chlTxnKey = ltTxn[4]
            chnlId = ltTxn[5]
            #补充相关信息
            txnMchtCd, txnTermId, amt, issAmt, swtAmt, prodAmt = getChnlTxn(db, chnlId, chlTxnKey)
            tailTxn04RptBody(ws, i, transDate, transTime, '长款', txnMchtCd, txnTermId,
                             '','',pan, amt, issAmt, swtAmt, prodAmt, amt - (issAmt+swtAmt+prodAmt))
        else :
            #后台成功
            keyRsp = ltTxn[3]
            mchtCd, termId, amt = getOwnTxn(db, keyRsp)
            tailTxn04RptBody(ws, i, transDate, transTime, '短款', '', '',
                             mchtCd, termId, pan, amt, 0.0, 0.0, 0.0, 0.0)
        i = i + 1

    cursor.close()

def tailTxn04RptBody(ws, i, transDate, transTime, type, txnMchtCd, txnTermId,
                     mchtCd, termId, pan, amt, issAmt, swtAmt, prodAmt, stlmAmt) :
    ws.cell(row=i, column=2).value = transDate
    ws.cell(row=i, column=3).value = transTime
    ws.cell(row=i, column=4).value = type
    ws.cell(row=i, column=5).value = txnMchtCd
    ws.cell(row=i, column=6).value = txnTermId
    ws.cell(row=i, column=7).value = mchtCd
    ws.cell(row=i, column=8).value = termId
    ws.cell(row=i, column=9).value = pan
    ws.cell(row=i, column=10).value = amt
    ws.cell(row=i, column=11).value = issAmt
    ws.cell(row=i, column=12).value = swtAmt
    ws.cell(row=i, column=13).value = prodAmt
    ws.cell(row=i, column=14).value = stlmAmt

#代付长短款交易明细
def handleTxn05RptHead(ws):
    global i
    i = i + 2
    ws.cell(row=i, column=1).value = '代付长短款交易明细'
    i = i + 1
    #报表头
    ws.cell(row=i, column=2).value = '交易日期'
    ws.cell(row=i, column=3).value = '类别'
    ws.cell(row=i, column=4).value = '代付商户号'
    ws.cell(row=i, column=5).value = '商户号'
    ws.cell(row=i, column=6).value = '终端号'
    ws.cell(row=i, column=7).value = '账户信息'
    ws.cell(row=i, column=8).value = '代付金额'
    ws.cell(row=i, column=9).value = '资金清算净额'
    i = i + 1

def tailTxn05RptBody(ws, i, transDate, type, txnMchtCd,
                     mchtCd, termId, pan, amt, stlmAmt) :
    ws.cell(row=i, column=2).value = transDate
    ws.cell(row=i, column=3).value = type
    ws.cell(row=i, column=4).value = txnMchtCd
    ws.cell(row=i, column=5).value = mchtCd
    ws.cell(row=i, column=6).value = termId
    ws.cell(row=i, column=7).value = pan
    ws.cell(row=i, column=8).value = amt
    ws.cell(row=i, column=9).value = stlmAmt


def handleTxn05RptBody(db, ws, stlm_date):
    #查找长短款交易
    global i
    sql = "select trans_date, CHK_STA, key_rsp, CHNL_RETRIVL_REF, GROUP_ID, " \
          " pan " \
          " from TBL_ERR_CHK_TXN_DTL " \
          "where host_date = '%s' and txn_num ='1801' and CHK_STA in ('1','2') order by  CHK_STA" % (stlm_date)
    cursor = db.cursor()
    cursor.execute(sql)
    for ltTxn in cursor:
        transDate = ltTxn[0]
        chkSta = ltTxn[1]
        pan = ltTxn[5]
        if chkSta == '1':
            chlTxnKey = ltTxn[3]
            chnlId = ltTxn[4]
            #补充相关信息
            txnMchtCd, txnTermId, amt, issAmt, swtAmt, prodAmt = getChnlTxn(db, chnlId, chlTxnKey)
            tailTxn05RptBody(ws, i, transDate, '短款', txnMchtCd,
                             '','',pan, amt, amt - issAmt - prodAmt)
        else :
            #后台成功
            keyRsp = ltTxn[2]
            mchtCd, termId, amt = getOwnTxn(db, keyRsp)
            tailTxn05RptBody(ws, i, transDate, '长款', '',
                             mchtCd, termId, pan, amt, 0.0)
        i = i + 1

    cursor.close()

def handleTxn05Rpt(db, ws, stlm_date):
    handleTxn05RptHead(ws)
    handleTxn05RptBody(db, ws, stlm_date)


def main():
    #数据库连接配置
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'],os.environ['DBPWD'],os.environ['TNSNAME']), encoding='gb18030')
    #获取清算日
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

    #生成指定文件
    wb = Workbook()
    ws = wb.active
    handleTxn01Rpt(db, ws, stlm_date)
    handleTxn02Rpt(db, ws, stlm_date)
    handleTxn03Rpt(db, ws, stlm_date)
    handleTxn04Rpt(db, ws, stlm_date)
    handleTxn05Rpt(db, ws, stlm_date)

    filePath = '%s/%s/' % (os.environ['RPT7HOME'],stlm_date)
    filename = filePath + 'AcqStlmCheckFile02_%s.xlsx' % stlm_date
    wb.save(filename)

if __name__ == '__main__':
    main()

