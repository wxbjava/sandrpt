#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
from utl.sndsms import sndsms

import cx_Oracle
import os


def main():
    # ���ݿ���������
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),
                           encoding='gb18030')
    cursor = db.cursor()
    sql = "select BF_STLM_DATE from TBL_BAT_CUT_CTL"
    cursor.execute(sql)
    x = cursor.fetchone()
    stlm_date = x[0]

    #��ȡ���������ܶ�
    sql = "select nvl(sum(REAL_TRANS_AMT),0) from TBL_STLM_TXN_BILL_DTL where host_date ='%s' and CHNL_ID ='A001'" % stlm_date
    cursor.execute(sql)
    x = cursor.fetchone()
    if x is None:
        transAmt = 0
    else:
        transAmt = x[0]

    #��ȡ����δƥ�佻�׽��
    sql = "select nvl(sum(REAL_TRANS_AMT),0) from TBL_STLM_TXN_BILL_DTL where stlm_date ='%s' and CHNL_ID ='A001' and check_sta ='0'" % stlm_date
    cursor.execute(sql)
    x = cursor.fetchone()
    if x is None:
        unknownAmt = 0
    else:
        unknownAmt = x[0]

    #��ȡ��ǰ�Ż��̻�������
    sql = "select count(*) from (select distinct mcht_cd from TBL_STLM_TXN_BILL_DTL where host_date ='%s' and CHNL_ID ='A001' and mcht_type ='2')" % stlm_date
    cursor.execute(sql)
    x = cursor.fetchone()
    if x is None:
        mchtCount = 0
    else:
        mchtCount = x[0]

    msg = "���ս����ܶ�:%.2fԪ,���򳤿��:%.2fԪ,�����Ż��̻�����:%d��" % (transAmt, unknownAmt, mchtCount)
    sms1 = sndsms()
    sms1.setPhone(['13917667716','17621110116'])
    sms1.setMsg(msg)
    sms1.sndSms()

    cursor.close()


if __name__ == '__main__':
    main()