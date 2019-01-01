#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#�����𱨱�,���ɱ����Լ���¼��ĩ������TBL_STLM_PVSN_RPT

import os
import sys
import cx_Oracle
from openpyxl.workbook import Workbook
from utl.common import *


class rptFile():
    def __init__(self, ws):
        self.ws = ws
        self.iCurr = 1

    def head(self):
        self.ws.cell(row=1, column=3).value = '�̻��������'
        self.ws.cell(row=1, column=4).value = 'δ������'
        self.ws.cell(row=1, column=5).value = '����������'
        self.ws.cell(row=1, column=6).value = 'Ӧ������'
        self.ws.cell(row=1, column=7).value = '����δ����'
        self.ws.cell(row=1, column=8).value = '�̻���֤��'
        self.ws.cell(row=1, column=9).value = '��ض���'
        self.ws.cell(row=1, column=10).value = '���д��'
        self.ws.cell(row=1, column=11).value = '���յ���'
        self.ws.cell(row=1, column=12).value = '��������'
        self.ws.cell(row=1, column=13).value = 'Ӧ��������������'
        self.iCurr = 2

    def recordInitAmt(self, initAcct):
        self.ws.cell(row=self.iCurr, column=1).value = '�ڳ�'
        self.ws.cell(row=self.iCurr, column=2).value = '����'
        self.ws.cell(row=self.iCurr, column=3).value = initAcct.mchtStlmAmt
        self.ws.cell(row=self.iCurr, column=4).value = initAcct.companyIncome
        self.ws.cell(row=self.iCurr, column=5).value = initAcct.insProfits
        self.ws.cell(row=self.iCurr, column=6).value = initAcct.chnlAmt
        self.ws.cell(row=self.iCurr, column=7).value = initAcct.diffAmt
        self.ws.cell(row=self.iCurr, column=8).value = initAcct.mchtDeposit
        self.ws.cell(row=self.iCurr, column=9).value = initAcct.lockAmt
        self.ws.cell(row=self.iCurr, column=10).value = initAcct.bankDeposit
        self.ws.cell(row=self.iCurr, column=11).value = initAcct.riskLoan
        self.ws.cell(row=self.iCurr, column=12).value = initAcct.othLoan
        self.ws.cell(row=self.iCurr, column=13).value = initAcct.payChnlLoan
        self.iCurr = self.iCurr + 1
    #���ս���
    def recordTodayTxn(self, initAcct, mchtStlmAmt, companyIncome,
                       insProfits, chnlAmt, diffAmt, riskLoan):
        self.ws.cell(row=self.iCurr, column=2).value = '���콻��'
        initAcct.mchtStlmAmt = toNumberFmt(initAcct.mchtStlmAmt + mchtStlmAmt)
        self.ws.cell(row=self.iCurr, column=3).value = mchtStlmAmt
        initAcct.companyIncome = toNumberFmt(initAcct.companyIncome + companyIncome)
        self.ws.cell(row=self.iCurr, column=4).value = companyIncome
        initAcct.insProfits = toNumberFmt(initAcct.insProfits + insProfits)
        self.ws.cell(row=self.iCurr, column=5).value = insProfits
        initAcct.chnlAmt = toNumberFmt(initAcct.chnlAmt + chnlAmt)
        self.ws.cell(row=self.iCurr, column=6).value = chnlAmt
        initAcct.diffAmt = toNumberFmt(initAcct.diffAmt + diffAmt)
        self.ws.cell(row=self.iCurr, column=7).value = diffAmt
        initAcct.riskLoan = toNumberFmt(initAcct.riskLoan + riskLoan)
        self.ws.cell(row=self.iCurr, column=11).value = riskLoan
        self.iCurr = self.iCurr + 1

    # ��ط��𶳽�
    def recordLockTxn(self, initAcct, lockAmt):
        self.ws.cell(row=self.iCurr, column=2).value = '��ط��𶳽�'
        initAcct.mchtStlmAmt = toNumberFmt(initAcct.mchtStlmAmt - lockAmt)
        self.ws.cell(row=self.iCurr, column=3).value = toNumberFmt(0 - lockAmt)
        initAcct.lockAmt = toNumberFmt(initAcct.lockAmt + lockAmt)
        self.ws.cell(row=self.iCurr, column=9).value = lockAmt
        self.iCurr = self.iCurr + 1

    def recordUnlockTxn(self, initAcct, unlockAmt):
        self.ws.cell(row=self.iCurr, column=2).value = '��ط���ⶳ'
        initAcct.mchtStlmAmt = toNumberFmt(initAcct.mchtStlmAmt + unlockAmt)
        self.ws.cell(row=self.iCurr, column=3).value = unlockAmt
        initAcct.lockAmt = toNumberFmt(initAcct.lockAmt - unlockAmt)
        self.ws.cell(row=self.iCurr, column=9).value = toNumberFmt(0 - unlockAmt)
        self.iCurr = self.iCurr + 1

    # ���
    def recordInAmt(self, initAcct, chnlAmt):
        self.ws.cell(row=self.iCurr, column=2).value = '���'
        initAcct.chnlAmt = toNumberFmt(initAcct.chnlAmt + chnlAmt)
        self.ws.cell(row=self.iCurr, column=6).value = initAcct.chnlAmt
        initAcct.bankDeposit = toNumberFmt(initAcct.bankDeposit - chnlAmt)
        self.ws.cell(row=self.iCurr, column=10).value = toNumberFmt(0 - chnlAmt)
        self.iCurr = self.iCurr + 1

    #�̻���������
    def recordMchtStlmAmtOut(self, initAcct, outMchtPayAmt):
        self.ws.cell(row=self.iCurr, column=2).value = '�̻���������'
        initAcct.mchtStlmAmt = toNumberFmt(initAcct.mchtStlmAmt - outMchtPayAmt)
        self.ws.cell(row=self.iCurr, column=3).value = toNumberFmt(0 - outMchtPayAmt)
        initAcct.payChnlLoan = toNumberFmt(initAcct.payChnlLoan + outMchtPayAmt)
        self.ws.cell(row=self.iCurr, column=13).value = outMchtPayAmt
        self.iCurr = self.iCurr + 1

    #�����������
    def recordInsIncomePayOut(self, initAcct, outAmt):
        self.ws.cell(row=self.iCurr, column=2).value = '�����������'
        initAcct.companyIncome = toNumberFmt(initAcct.companyIncome - outAmt)
        self.ws.cell(row=self.iCurr, column=4).value = toNumberFmt(0 - outAmt)
        initAcct.payChnlLoan = toNumberFmt(initAcct.payChnlLoan + outAmt)
        self.ws.cell(row=self.iCurr, column=13).value = outAmt
        self.iCurr = self.iCurr + 1

    #�ʽ������۴���
    def recordChnlPayAmt(self, initAcct, chnlPayAmt):
        self.ws.cell(row=self.iCurr, column=2).value = '�ʽ������۴���'
        initAcct.bankDeposit = toNumberFmt(initAcct.bankDeposit + chnlPayAmt)
        self.ws.cell(row=self.iCurr, column=10).value = chnlPayAmt
        initAcct.payChnlLoan = toNumberFmt(initAcct.payChnlLoan - chnlPayAmt)
        self.ws.cell(row=self.iCurr, column=13).value = toNumberFmt(0 - chnlPayAmt)
        self.iCurr = self.iCurr + 1

    def recordFinalAmt(self, initAcct):
        self.ws.cell(row=self.iCurr, column=1).value = '��ĩ'

        self.ws.cell(row=self.iCurr, column=3).value = initAcct.mchtStlmAmt
        self.ws.cell(row=self.iCurr, column=4).value = initAcct.companyIncome
        self.ws.cell(row=self.iCurr, column=5).value = initAcct.insProfits
        self.ws.cell(row=self.iCurr, column=6).value = initAcct.chnlAmt
        self.ws.cell(row=self.iCurr, column=7).value = initAcct.diffAmt
        self.ws.cell(row=self.iCurr, column=8).value = initAcct.mchtDeposit
        self.ws.cell(row=self.iCurr, column=9).value = initAcct.lockAmt
        self.ws.cell(row=self.iCurr, column=10).value = initAcct.bankDeposit
        self.ws.cell(row=self.iCurr, column=11).value = initAcct.riskLoan
        self.ws.cell(row=self.iCurr, column=12).value = initAcct.othLoan
        self.ws.cell(row=self.iCurr, column=13).value = initAcct.payChnlLoan
        self.iCurr = self.iCurr + 1

class stlmPvsnAcctInfo:
    def __init__(self, db, stlmDate):
        self.db = db
        self.stlmDate = stlmDate
        self.mchtStlmAmt = 0
        self.companyIncome = 0
        self.insProfits = 0
        self.chnlAmt = 0
        self.diffAmt = 0
        self.mchtDeposit = 0
        self.lockAmt = 0
        self.bankDeposit = 0
        self.riskLoan = 0
        self.othLoan = 0
        self.payChnlLoan = 0
        self.__get_init_acct_info()

    #��ȡ������ĩ
    def __get_init_acct_info(self):
        sql = "select MCHT_STLM_AMT,COMPANY_INCOME,INS_PROFITS,CHNL_AMT, " \
              "DIFF_AMT,MCHT_DEPOSIT,LOCK_AMT,BANK_DEPOSIT, " \
              "RISK_LOAN,OTH_LOAN,PAY_CHNL_LOAN  from TBL_STLM_PVSN_RPT where host_date ='%s'" % getLastDay(self.stlmDate)
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x is not None:
            self.mchtStlmAmt = toNumberFmt(x[0])
            self.companyIncome = toNumberFmt(x[1])
            self.insProfits = toNumberFmt(x[2])
            self.chnlAmt = toNumberFmt(x[3])
            self.diffAmt = toNumberFmt(x[4])
            self.mchtDeposit = toNumberFmt(x[5])
            self.lockAmt = toNumberFmt(x[6])
            self.bankDeposit = toNumberFmt(x[7])
            self.riskLoan = toNumberFmt(x[8])
            self.othLoan = toNumberFmt(x[9])
            self.payChnlLoan = toNumberFmt(x[10])

    #��ֵ��¼���ݿ�
    def insertFinalInfo(self):
        sql = "insert into TBL_STLM_PVSN_RPT "
        sqltmp = "(host_date,MCHT_STLM_AMT,COMPANY_INCOME,INS_PROFITS,CHNL_AMT, " \
                 "DIFF_AMT,MCHT_DEPOSIT,LOCK_AMT,BANK_DEPOSIT, " \
                 "RISK_LOAN,OTH_LOAN,PAY_CHNL_LOAN ) values "
        sql = sql + sqltmp
        sqltmp = "('%s',%f, %f, %f, %f, %f, %f, %f," \
                 "%f, %f, %f, %f)" % \
                 (self.stlmDate, self.mchtStlmAmt, self.companyIncome, self.insProfits, self.chnlAmt,
                  self.diffAmt, self.mchtDeposit, self.lockAmt, self.bankDeposit, self.riskLoan,
                  self.othLoan, self.payChnlLoan)
        sql = sql + sqltmp
        cursor = self.db.cursor()
        print(sql)
        cursor.execute(sql)
        self.db.commit()
        cursor.close()

#���������ֵͳ��
class txnInfo:
    def __init__(self, db, stlmDate):
        self.db = db
        self.stlmDate = stlmDate
        self.mchtStlmAmt = 0
        self.companyIncome = 0
        self.insIncome = 0
        self.diffChnlAmt = 0
        self.riskLoan = 0
        self.mchtPayAmt = 0
        self.agentPayAmt = 0

        self.__get_mcht_stlm()
        self.__get_company_income()
        self.__get_ins_income()
        self.__get_diff_chnl_amt()
        self.__get_risk_loan()
        self.__get_paytxn_amt()

        self.calcChnlAmt = toNumberFmt(0 - self.mchtStlmAmt - self.companyIncome - self.insIncome - self.diffChnlAmt - self.riskLoan)

    #����������
    def __get_mcht_stlm(self):
        sql = "select sum(TRANS_AMT - TRANS_FEE) from TBL_INS_PROFITS_TXN_SUM where " \
              "host_date ='%s' and PROFITS_TYPE ='00'" % self.stlmDate
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x[0] is not None:
            self.mchtStlmAmt = toNumberFmt(x[0])
        cursor.close()

    #��˾δ������(��˾����)
    def __get_company_income(self):
        sql = "select sum(ALL_PROFITS) from TBL_SAND_ACQ_PROFITS where " \
              "host_date ='%s'" % self.stlmDate
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x[0] is not None:
            self.companyIncome = toNumberFmt(x[0])
        cursor.close()

    #��������������
    def __get_ins_income(self):
        sql = "select sum(ALL_PROFITS) from TBL_INS_PROFITS_TXN_SUM where " \
              "host_date ='%s'" % self.stlmDate
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x[0] is not None:
            self.insIncome = toNumberFmt(x[0])
        cursor.close()

    #����δ����(���ճ��� - ���ճ��� + ����)
    def __get_diff_chnl_amt(self):
        sql = "select sum(REAL_TRANS_AMT) from tbl_stlm_txn_bill_dtl where host_date ='%s' and " \
              "stlm_date = '%s' and check_sta ='1' and txn_num ='1011'" \
              % (self.stlmDate, getLastDay(self.stlmDate))
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x is not None:
            lastAmt = toNumberFmt(x[0])
        else:
            lastAmt = 0
        sql = "select sum(REAL_TRANS_AMT) from tbl_stlm_txn_bill_dtl where host_date ='%s' and " \
              "stlm_date = '%s' and check_sta ='1' and txn_num ='1011'" \
              % (getNextDay(self.stlmDate), self.stlmDate)
        cursor.execute(sql)
        x = cursor.fetchone()
        if x is not None:
            currAmt = toNumberFmt(x[0])
        else:
            currAmt = 0

        sql = "select sum(CHNL_TXN_AMT) from tbl_err_chk_txn_dtl where " \
              "host_date = '%s' and CHK_STA ='1' and group_id ='A001'" % self.stlmDate
        cursor.execute(sql)
        x = cursor.fetchone()
        if x is not None:
            longAmt = toNumberFmt(x[0])
        else:
            longAmt = 0
        self.diffChnlAmt = longAmt + currAmt - lastAmt
        cursor.close()

    #���յ���
    def __get_risk_loan(self):
        #��ʱ��֪����������,��ʱΪ��
        self.riskLoan = 0

    #�ɹ��������
    def __get_paytxn_amt(self):
        sql = "select sum(REAL_TRANS_AMT) from tbl_stlm_txn_bill_dtl where " \
              "chnl_id ='A002' and host_date ='%s'" % self.stlmDate
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x[0] is not None:
            self.mchtPayAmt = toNumberFmt(0 - x[0])
        #����������
        sql = "select sum(trans_amt/100) from tbl_acq_txn_log where host_date ='%s' and " \
              "txn_num ='1801' and substrb(ADDTNL_DATA,1,2) ='04' and trans_state ='1'" % self.stlmDate
        cursor.execute(sql)
        x = cursor.fetchone()
        if x[0] is not None:
            self.agentPayAmt = toNumberFmt(x[0])
            self.mchtPayAmt = toNumberFmt(self.mchtPayAmt - self.agentPayAmt)
        cursor.close()



class lockInfo:
    def __init__(self, db, stlmDate):
        self.db = db
        self.stlmDate = stlmDate
        self.lockAmt = 0
        self.unlockAmt = 0

        self.__get_lock_amt()

    def __get_lock_amt(self):
        sql = "select sum(LOCK_AT)/100 from T_TXN_LOCK where " \
              "host_date ='%s' and TXN_TYPE ='01' " % self.stlmDate
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x[0] is not None:
            self.lockAmt = toNumberFmt(x[0])
        cursor.close()

    def __get_unlock_amt(self):
        sql = "select sum(free_at)/100 from T_TXN_LOCK where" \
              "host_date ='%s' and TXN_TYPE ='02' " % self.stlmDate
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x[0] is not None:
            self.unlockAmt = toNumberFmt(x[0])
        cursor.close()

class chnlAmtInfo:
    def __init__(self, db, stlmDate):
        self.db = db
        self.stlmDate = stlmDate
        self.intxnAmt = 0
        self.outtxnAmt = 0

        self.__get_intxn_amt()
        self.__get_outtxn_amt()

    #�������ļ�
    def __get_intxn_amt(self):
        sql = "select sum(REAL_TRANS_AMT) from tbl_stlm_txn_bill_dtl where chnl_id ='A001' and " \
              "stlm_date ='%s'" % self.stlmDate
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x[0] is not None:
            self.intxnAmt = toNumberFmt(x[0])
        cursor.close()


    #�����ļ�
    def __get_outtxn_amt(self):
        sql = "select sum(REAL_TRANS_AMT) from tbl_stlm_txn_bill_dtl where chnl_id ='A002' and " \
              "stlm_date ='%s'" % self.stlmDate
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x[0] is not None:
            self.outtxnAmt = toNumberFmt(0 - x[0])
        cursor.close()



def main():
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),
                           encoding='gb18030')
    dbacc = cx_Oracle.connect('%s/%s@%s' % (os.environ['ACCDBUSER'], os.environ['ACCDBPWD'], os.environ['TNSNAME']),
                           encoding='gb18030')
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
    print('hostDate %s genStlmPvsnRpt begin' % stlm_date)

    filePath = '%s/%s/' % (os.environ['RPT7HOME'], stlm_date)
    filename = filePath + 'StlmPvsnRpt_%s.xlsx' % stlm_date

    stlmPvsnAcct = stlmPvsnAcctInfo(db, stlm_date)
    wb = Workbook()
    ws = wb.active
    rptxls = rptFile(ws)
    rptxls.head()
    rptxls.recordInitAmt(stlmPvsnAcct)
    #���콻��
    txnInf = txnInfo(db, stlm_date)
    rptxls.recordTodayTxn(stlmPvsnAcct, txnInf.mchtStlmAmt, txnInf.companyIncome,
                          txnInf.insIncome, txnInf.calcChnlAmt, txnInf.diffChnlAmt,
                          toNumberFmt(0 - txnInf.riskLoan))
    #��ط��𶳽�
    lockInf = lockInfo(dbacc, stlm_date)
    rptxls.recordLockTxn(stlmPvsnAcct, lockInf.lockAmt)
    #��ط���ⶳ
    rptxls.recordUnlockTxn(stlmPvsnAcct, lockInf.unlockAmt)
    #���
    chnlFile = chnlAmtInfo(db, stlm_date)
    rptxls.recordInAmt(stlmPvsnAcct, chnlFile.intxnAmt)
    #�̻���������
    rptxls.recordMchtStlmAmtOut(stlmPvsnAcct, txnInf.mchtPayAmt)
    #�����(����)

    #�����������
    rptxls.recordInsIncomePayOut(stlmPvsnAcct, txnInf.agentPayAmt)
    #�ʽ������۴���
    rptxls.recordChnlPayAmt(stlmPvsnAcct, chnlFile.outtxnAmt)

    #��ĩ
    rptxls.recordFinalAmt(stlmPvsnAcct)

    stlmPvsnAcct.insertFinalInfo()

    wb.save(filename)
    wb.close()
    print('hostDate %s genStlmPvsnRpt end' % stlm_date)

if __name__ == '__main__':
    main()