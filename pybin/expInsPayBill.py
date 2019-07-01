#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#机构分润出款对账单

import os
import sys
import cx_Oracle
from openpyxl.workbook import Workbook
from utl.common import *

class rptFile():
    def __init__(self, ins_id_cd = None, stlm_date = None):
        self.iCurr = 1
        self.ins_id_cd = ins_id_cd
        self.stlm_date = stlm_date
        self.wb = None
        if self.ins_id_cd != None:
            self.wb = Workbook()
            self.ws = self.wb.active
            self.__fileHeader()

    def __fileHeader(self):
        self.ws.cell(row=self.iCurr, column=1).value = '日期'
        self.ws.cell(row=self.iCurr, column=2).value = '时间'
        self.ws.cell(row=self.iCurr, column=3).value = '代付订单号'
        self.ws.cell(row=self.iCurr, column=4).value = '商户订单号'
        self.ws.cell(row=self.iCurr, column=5).value = '交易类型'
        self.ws.cell(row=self.iCurr, column=6).value = '发起方式'
        self.ws.cell(row=self.iCurr, column=7).value = '代付金额'
        self.ws.cell(row=self.iCurr, column=8).value = '日期'

        self.iCurr = self.iCurr + 1


    def getInsId(self):
        return self.ins_id_cd

    def saveFile(self):
        if self.wb != None:
            filePath = '%s/%s/' % (os.environ['RPT7HOME'], self.stlm_date)
            self.wb.save('%s/InsPayBill_%s.xlsx'% (filePath, self.stlm_date))
            self.wb.close()







def main():
    # 数据库连接配置
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
    print('hostDate %s rpt begin' % stlm_date)


    #查找指定日期机构分润代付情况
    sql = "select a.key_rsp, a.TXN_NUM, a.TXN_DATE, a.TXN_TIME, trim(a.INS_ID_CD), trim(substrb(b.ADDTNL_DATA, 3, 28))," \
          "trim(substrb(b.ADDTNL_DATA, 31, 60)), trim(substrb(b.ADDTNL_DATA, 91, 12)), trim(substrb(b.ADDTNL_DATA, 103, 60)), " \
          "b.trans_amt / 100, b.next_txn_key, b.MSQ_TYPE" \
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



    print('hostDate %s rpt end' % stlm_date)

if __name__ == '__main__':
    main()