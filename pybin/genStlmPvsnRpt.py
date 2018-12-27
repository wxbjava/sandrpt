#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#备付金报表,生成报表以及记录期末数据于TBL_STLM_PVSN_RPT

import os
import sys
import cx_Oracle
from openpyxl.workbook import Workbook
from utl.common import *



def main():
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),
                           encoding='gb18030')


if __name__ == '__main__':
    main()