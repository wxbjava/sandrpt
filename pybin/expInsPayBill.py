#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#�������������˵�

import os
import sys
import cx_Oracle
from openpyxl.workbook import Workbook
from utl.common import *


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
    sql = "select a.key_rsp, a.TXN_NUM, a.TXN_DATE, a.TXN_TIME, a.INS_ID_CD, "

    print('hostDate %s rpt end' % stlm_date)

if __name__ == '__main__':
    main()