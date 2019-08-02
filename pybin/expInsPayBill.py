#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#�������������˵�

import os
import sys
import cx_Oracle
from openpyxl.workbook import Workbook
from utl.common import *

class rptFile():
    def __init__(self, ins_id_cd=None, stlm_date=None):
        self.ins_id_cd = ins_id_cd
        self.stlm_date = stlm_date
        self.wb = None
        if self.ins_id_cd != None:
            self.wb = Workbook(write_only=True)
            self.ws = self.wb.create_sheet()
            self.__fileHeader()

    def __fileHeader(self):
        data = []
        data.append('����')
        data.append('ʱ��')
        data.append('����������')
        data.append('�̻�������')
        data.append('��������')
        data.append('����ʽ')
        data.append('�������')
        data.append('���㿨')
        data.append('�����˻���')
        data.append('���к�')
        data.append('��������')
        data.append('���ױ�ע')
        self.ws.append(data)


    def tailData(self, instDate, instTime, payOrder, reqOrder, txnName, txnType,
                 payAmt, acctNo, acctNm, bankId, bankNm, payDesc):
        data = []
        data.append(instDate)
        data.append(instTime)
        data.append(payOrder)
        data.append(reqOrder)
        data.append(txnName)
        data.append(txnType)
        data.append(payAmt)
        data.append(acctNo)
        data.append(acctNm)
        data.append(bankId)
        data.append(bankNm)
        data.append(payDesc)
        self.ws.append(data)


    def getInsId(self):
        return self.ins_id_cd

    def saveFile(self):
        if self.wb != None:
            filePath = '%s/%s/' % (os.environ['RPT7HOME'], self.stlm_date)
            self.wb.save('%s/InsPayBill_%s_%s.xlsx'% (filePath, self.ins_id_cd, self.stlm_date))
            self.wb.close()

def getDesc(db, keyRsp):
    sql = "select substrb(key_cancel,1,16) from tbl_acq_txn_log where key_rsp = '%s'" % keyRsp
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    if x is not None:
        sql = "select trim(PAY_DESC) from TBL_DEST_PAY_LOG where key_rsp = '%s'" % x[0]
        cursor.execute(sql)
        x = cursor.fetchone()
        if x is not None:
            return x[0]

    return " "


def main():
    # ���ݿ���������
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
    print('hostDate %s rpt begin' % stlm_date)


    #����ָ�����ڻ�������������
    sql = "select a.key_rsp, a.TXN_NUM, a.TXN_DATE, a.TXN_TIME, trim(a.INS_ID_CD), trim(substrb(b.ADDTNL_DATA, 3, 28))," \
          "trim(substrb(b.ADDTNL_DATA, 31, 60)), trim(substrb(b.ADDTNL_DATA, 91, 12)), trim(substrb(b.ADDTNL_DATA, 103, 60)), " \
          "b.trans_amt / 100, b.next_txn_key, trim(b.MSQ_TYPE)" \
          "from TBL_STLM_TXN_BILL_DTL a left join tbl_acq_txn_log b on a.key_rsp = b.key_rsp " \
          "where a.host_date ='%s' and a.PAY_TYPE ='03' order by a.ins_id_cd, a.txn_date, a.txn_time " % stlm_date
    cursor = db.cursor()
    cursor.execute(sql)
    rptfile = rptFile()
    for ltData in cursor:
        ins_id_cd = ltData[4]
        if ins_id_cd != rptfile.getInsId():
            rptfile.saveFile()
            rptfile = rptFile(ins_id_cd = ins_id_cd, stlm_date = stlm_date)
        instDate = ltData[2]
        instTime = ltData[3]
        payOrder = ltData[0]
        reqOrder = ltData[10]
        payDesc = ''
        if ltData[1] == '1801':
            txnName = '����'
        elif ltData[1] == '9164':
            txnName = '�˵�'
        else:
            txnName = 'δ֪����'
        if ltData[11] == '-2':
            payDesc = getDesc(db, payOrder)
            txnType = 'ƽ̨����'
        else:
            txnType = '�ӿڴ���'
        payAmt = ltData[9]
        acctNo = ltData[5]
        acctNm = ltData[6]
        bankId = ltData[7]
        bankNm = ltData[8]
        rptfile.tailData(instDate, instTime, payOrder, reqOrder, txnName, txnType,payAmt, acctNo, acctNm, bankId, bankNm, payDesc)

    rptfile.saveFile()

    print('hostDate %s rpt end' % stlm_date)

if __name__ == '__main__':
    main()