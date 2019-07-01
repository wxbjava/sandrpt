#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#机构分润出款对账单

import os
import sys
import cx_Oracle
from openpyxl.workbook import Workbook
from utl.common import *


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
    sql = "select a.key_rsp, a.TXN_NUM, a.TXN_DATE, a.TXN_TIME, a.INS_ID_CD, "

    print('hostDate %s rpt end' % stlm_date)

if __name__ == '__main__':
    main()