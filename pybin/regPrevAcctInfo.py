#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#按机构留存上日余额,供余额报表使用

import cx_Oracle
import os
import sys

class sandBalance:
    def __init__(self, insId, stlmDate, dbacc, dbbat):
        self.insId = insId
        self.stlmDate = stlmDate
        self.dbacc = dbacc
        self.dbbat = dbbat

    def __get_mcht_prev_at(self):
        #待结算
        sql = "select sum(case when a.prev_bal_dt <= '%s' then a.curr_bal_at else a.prev_bal_at end), " \
              "sum(case when a.prev_avail_dt <= '%s' then a.curr_avail_at else a.prev_avail_at end) " \
          " from (select * from  t_acct_info where ACCT_TYPE ='00000001') a left join" \
          " t_acct_map b on a.acct_id = b.acct_id left join" \
          " tbl_mcht_inf c on b.ext_acct_id = c.mcht_cd where c.company_cd ='%s'" % (self.stlmDate, self.stlmDate, self.insId)
        cursor = self.dbacc.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        cursor.close()
        if x[0] is not None:
            self.mcht_a_prev_bal_at = x[0]
            self.mcht_a_prev_avail_at = x[1]
        else:
            self.mcht_a_prev_bal_at = 0
            self.mcht_a_prev_avail_at = 0

        #结算
        sql = "select sum(case when a.prev_bal_dt <= '%s' then a.curr_bal_at else a.prev_bal_at end), " \
              "sum(case when a.prev_avail_dt <= '%s' then a.curr_avail_at else a.prev_avail_at end) " \
              " from (select * from  t_acct_info where ACCT_TYPE ='00000002') a left join" \
              " t_acct_map b on a.acct_id = b.acct_id left join" \
              " tbl_mcht_inf c on b.ext_acct_id = c.mcht_cd where c.company_cd ='%s'" % (
              self.stlmDate, self.stlmDate, self.insId)
        cursor = self.dbacc.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        cursor.close()
        if x[0] is not None:
            self.mcht_b_prev_bal_at = x[0]
            self.mcht_b_prev_avail_at = x[1]
        else:
            self.mcht_b_prev_bal_at = 0
            self.mcht_b_prev_avail_at = 0

        #欠款
        sql = "select sum(case when a.prev_bal_dt <= '%s' then a.curr_bal_at else a.prev_bal_at end), " \
              "sum(case when a.prev_avail_dt <= '%s' then a.curr_avail_at else a.prev_avail_at end) " \
              " from (select * from  t_acct_info where ACCT_TYPE ='00000003') a left join" \
              " t_acct_map b on a.acct_id = b.acct_id left join" \
              " tbl_mcht_inf c on b.ext_acct_id = c.mcht_cd where c.company_cd ='%s'" % (
                  self.stlmDate, self.stlmDate, self.insId)
        cursor = self.dbacc.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        cursor.close()
        if x[0] is not None:
            self.mcht_c_prev_bal_at = x[0]
            self.mcht_c_prev_avail_at = x[1]
        else:
            self.mcht_c_prev_bal_at = 0
            self.mcht_c_prev_avail_at = 0

    def __get_ins_prev_at(self):
        #分润户
        sql = "select sum(case when a.prev_bal_dt <= '%s' then a.curr_bal_at else a.prev_bal_at end), " \
              "sum(case when a.prev_avail_dt <= '%s' then a.curr_avail_at else a.prev_avail_at end) " \
              " from (select * from  t_acct_info where ACCT_TYPE ='00000002') a left join" \
              " t_acct_map b on a.acct_id = b.acct_id " \
              "where b.ext_acct_id ='%sA'" % (
                  self.stlmDate, self.stlmDate, self.insId)
        cursor = self.dbacc.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        cursor.close()
        if x[0] is not None:
            self.ins_b_prev_bal_at = x[0]
            self.ins_b_prev_avail_at = x[1]
        else:
            self.ins_b_prev_bal_at = 0
            self.ins_b_prev_avail_at = 0

        #欠款
        sql = "select sum(case when a.prev_bal_dt <= '%s' then a.curr_bal_at else a.prev_bal_at end), " \
              "sum(case when a.prev_avail_dt <= '%s' then a.curr_avail_at else a.prev_avail_at end) " \
              " from (select * from  t_acct_info where ACCT_TYPE ='00000003') a left join" \
              " t_acct_map b on a.acct_id = b.acct_id " \
              "where b.ext_acct_id ='%sA'" % (
                  self.stlmDate, self.stlmDate, self.insId)
        cursor = self.dbacc.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        cursor.close()
        if x[0] is not None:
            self.ins_c_prev_bal_at = x[0]
            self.ins_c_prev_avail_at = x[1]
        else:
            self.ins_c_prev_bal_at = 0
            self.ins_c_prev_avail_at = 0

    def __get_acq_prev_at(self):
        sql = "select sum(case when a.prev_bal_dt <= '%s' then a.curr_bal_at else a.prev_bal_at end), " \
              "sum(case when a.prev_avail_dt <= '%s' then a.curr_avail_at else a.prev_avail_at end) " \
              " from (select * from  t_acct_info where ACCT_TYPE ='00000002') a left join" \
              " t_acct_map b on a.acct_id = b.acct_id " \
              "where b.ext_acct_id ='%sB'" % (
                  self.stlmDate, self.stlmDate, self.insId)
        cursor = self.dbacc.cursor()
        cursor.execute(sql)
        x = cursor.fetchone()
        cursor.close()
        if x[0] is not None:
            self.acq_prev_bal_at = x[0]
            self.acq_prev_avail_at = x[1]
        else:
            self.acq_prev_bal_at = 0
            self.acq_prev_avail_at = 0

    def insertDb(self):
        self.__get_mcht_prev_at()
        self.__get_ins_prev_at()
        self.__get_acq_prev_at()

        sqltmp = "insert into TBL_SAND_BALANCE_INF "
        sql = sqltmp
        sqltmp = "(HOST_DATE,INS_ID_CD,MCHT_A_PREV_BAL_AT,MCHT_A_PREV_AVAIL_AT," \
                 "MCHT_B_PREV_BAL_AT,MCHT_B_PREV_AVAIL_AT,MCHT_C_PREV_BAL_AT,MCHT_C_PREV_AVAIL_AT," \
                 "INS_B_PREV_BAL_AT,INS_B_PREV_AVAIL_AT,INS_C_PREV_BAL_AT,INS_C_PREV_AVAIL_AT,ACQ_PREV_BAL_AT,ACQ_PREV_AVAIL_AT) " \
                 "values "
        sql = sql + sqltmp
        sqltmp = "('%s','%s',%d,%d,%d,%d," \
                 "%d, %d, %d, %d," \
                 "%d, %d, %d, %d)" % (self.stlmDate, self.insId,
                    self.mcht_a_prev_bal_at, self.mcht_a_prev_avail_at,
                    self.mcht_b_prev_bal_at, self.mcht_b_prev_avail_at,
                    self.mcht_c_prev_bal_at, self.mcht_c_prev_avail_at,
                    self.ins_b_prev_bal_at, self.ins_b_prev_avail_at,
                    self.ins_c_prev_bal_at, self.ins_c_prev_avail_at,
                    self.acq_prev_bal_at, self.acq_prev_avail_at)
        sql = sql + sqltmp
        print(sql)
        cursor = self.dbbat.cursor()
        cursor.execute(sql)
        self.dbbat.commit()
        cursor.close()


def main():
    # 数据库连接配置
    dbbat = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),
                              encoding='gb18030')
    dbacc = cx_Oracle.connect('%s/%s@%s' % (os.environ['ACCDBUSER'], os.environ['ACCDBPWD'], os.environ['TNSNAME']),
                              encoding='gb18030')

    # 获取清算日
    if len(sys.argv) == 1:
        cursor = dbbat.cursor()
        sql = "select BF_STLM_DATE from TBL_BAT_CUT_CTL"
        cursor.execute(sql)
        x = cursor.fetchone()
        stlm_date = x[0]
        cursor.close()
    else:
        stlm_date = sys.argv[1]

    print('hostDate %s genRptAcqBalance begin' % stlm_date)
    #查找机构
    sql = "select trim(INS_ID_CD) from TBL_INS_INF where INS_TP ='01'"
    cursor = dbbat.cursor()
    cursor.execute(sql)
    for ltData in cursor:
        if ltData[0] is not None:
            sandBal = sandBalance(ltData[0], stlm_date, dbacc, dbbat)
            sandBal.insertDb()

    cursor.close()

if __name__ == '__main__':
    main()