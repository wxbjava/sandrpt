#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-

#收益核销

import cx_Oracle
import sys
import hashlib
import os
from utl.common import *


def getAcqIncome(db, stlmDate):
    #杉德收入
    sql = "select sum(ALL_PROFITS) " \
          "from TBL_SAND_ACQ_PROFITS where host_date = '%s' " % stlmDate
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    if x is not None:
        acqIncome = toNumberFmt(x[0])
    else:
        acqIncome = 0
    return acqIncome

def get_file_sha1(f):
    m = hashlib.sha1()
    while True:
        data = f.read(10240)
        if not data:
            break

        m.update(data)
    return m.hexdigest().upper()

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

    filePath = '%s/%s/' % (os.environ['RPT7HOME'], stlm_date)
    filename = filePath + '1005600100000000%s0001.TXT' % getNextDay(stlm_date)
    fd = open(filename,"w", encoding="gb18030")
    amt = getAcqIncome(db, stlm_date)

    now = datetime.datetime.now()
    txnline = "%s0-00000000000001|%s|%s|%s|%s|200250  |200250              |" \
              "48270000|00000400|00000400|9007|156|0     |%-12.2f|\n" % (now.strftime('%Y%m%d%H%M%S'),
                                                  stlm_date,getNextDay(stlm_date),
                                                  stlm_date, getNextDay(stlm_date), amt)
    print(txnline)

    headline = '%08d|1234567890123456789012345678901234567890|1.0.0|1005|0|00000000|%s|0001|%.2f\n' % (1,stlm_date,amt)
    fd.write(headline)
    fd.write(txnline)
    fd.close()

    #打开文件
    fd = open(filename, "rb")
    #计算sha1
    sha = get_file_sha1(fd)
    fd.close()

    fd = open(filename,"r+", encoding="gb18030")
    fd.seek(9,0)
    fd.write(sha)
    fd.close()



if __name__ == '__main__':
    main()