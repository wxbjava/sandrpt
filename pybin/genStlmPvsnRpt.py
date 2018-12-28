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
            self.mchtStlmAmt = x[0]
            self.companyIncome = x[1]
            self.insProfits = x[2]
            self.chnlAmt = x[3]
            self.diffAmt = x[4]
            self.mchtDeposit = x[5]
            self.lockAmt = x[6]
            self.bankDeposit = x[7]
            self.riskLoan = x[8]
            self.othLoan = x[9]
            self.payChnlLoan = x[10]

#���������ֵͳ��
class TxnInfo:
    def __init__(self, db, stlmDate):
        self.db = db
        self.stlmDate = stlmDate
        self.mchtStlmAmt = 0
        self.companyIncome = 0
        self.insIncome = 0


        self.__get_mcht_stlm()
        self.__get_company_income()
        self.__get_ins_income()

    #����������
    def __get_mcht_stlm(self):
        sql = "select sum(TRANS_AMT - TRANS_FEE) from TBL_INS_PROFITS_TXN_SUM where " \
              "host_date ='%s' and PROFITS_TYPE ='00'" % self.stlmDate
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x[0] is not None:
            self.mchtStlmAmt = x[0]
        cursor.close()

    #��˾δ������(��˾����)
    def __get_company_income(self):
        sql = "select sum(ALL_PROFITS) from TBL_SAND_ACQ_PROFITS where " \
              "host_date ='%s'" % self.stlmDate
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x[0] is not None:
            self.companyIncome = x[0]
        cursor.close()

    #��������������
    def __get_ins_income(self):
        sql = "select sum(ALL_PROFITS) from TBL_INS_PROFITS_TXN_SUM where" \
              "host_date ='%s'" % self.stlmDate
        cursor = self.db.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        if x[0] is not None:
            self.insIncome = x[0]
        cursor.close()

    #����δ����(���ճ��� - ���ճ��� + ����)




def main():
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),
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




    wb.save(filename)
    wb.close()
    print('hostDate %s genStlmPvsnRpt end' % stlm_date)

if __name__ == '__main__':
    main()