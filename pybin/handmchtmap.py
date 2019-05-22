#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-

import xlrd
import cx_Oracle
import os

bin_id = '9892'
db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSERSWT'], os.environ['DBPWDSWT'], os.environ['TNSNAME']),
                           encoding='gb18030')

filename = "E:\\需求\\杉德系统数据变更\\商户池\\20190520第十批标扣/data20190520.sql"
fin = open(filename, "wb")


unsportMccList = [
'5411',
'4511',
'4121',
'5541',
'5542',
'1520',
'4011',
'4111',
'4119',
'4131',
'4784',
'4789',
'4900',
'5933',
'5935',
'5960',
'5976',
'6211',
'6300',
'7261',
'7273',
'7276',
'7277',
'7321',
'7995',
'8011',
'8031',
'8049',
'8050',
'8062',
'8099',
'8111',
'8211',
'8220',
'8398',
'8641',
'8651',
'8661',
'9211',
'9222',
'9223',
'9311',
'9399',
'9400',
'9402',
'6010',
'6011',
'6012',
'6051',
'7013',
'9498',
'9708'
]

def getMccInfo(mcc_cd):
    sql = "select HIGH_AMT, BEGIN_TIME, END_TIME from TBL_SHC_MCC_CFG where MCC_CD ='%s'" % mcc_cd
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    if x is not None:
        return x[0],x[1],x[2]
    else:
        print("未查到 %s" % mcc_cd)
        return '*','*','*'

def genSql(mcc_cd, obj_mcht_cd, obj_term_id, mcht_name, area_cd, acq_ins_id):
    sql = "INSERT INTO TBL_SHC_MCHT_TERM_MAP (MAP_IDX, BIN_ID, BANK_ID, CARD_TP, CARD_BIN, MCC_TP, MCHT_GP, SRC_MCHT_CD, SRC_TERM_ID, SRC_MCHT_TP, SRC_SPEC_FEE_CD, LOW_AMT, HIGH_AMT, START_DATE, START_TIME, END_DATE, END_TIME, ACQ_INS_ID_CD, FWD_INS_ID_CD, OBJ_MCHT_CD, OBJ_TERM_ID, OBJ_MCHT_TP, OBJ_MCHT_NAME, OBJ_MCHT_NAME_OTH, OBJ_ACQ_INS_ID_CD, OBJ_SPEC_FEE_CD, REC_OPR_ID, REC_UPD_OPR, REC_CRT_TS, REC_UPD_TS, AREA_CD) VALUES " \
          "(5, '%s', '*', '* ', '*', '*', '*', '*', '*', '%s', '*  ', '*', '*', '*', '*', '*', '*', '09970000', '*          ', '%s', '%s', '%s', '%s','%s', '%s', '*  ', '*', '*         ', sysdate, sysdate, '%s');\n" % (bin_id, mcc_cd, obj_mcht_cd, obj_term_id,mcc_cd, mcht_name, mcht_name, acq_ins_id, area_cd)
    fin.write(sql.encode('gb18030'))

    if mcc_cd not in unsportMccList:
        high_amt, start_time, end_time = getMccInfo(mcc_cd)
        sql = "INSERT INTO TBL_SHC_MCHT_TERM_MAP (MAP_IDX, BIN_ID, BANK_ID, CARD_TP, CARD_BIN, MCC_TP, MCHT_GP, SRC_MCHT_CD, SRC_TERM_ID, SRC_MCHT_TP, SRC_SPEC_FEE_CD, LOW_AMT, HIGH_AMT, START_DATE, START_TIME, END_DATE, END_TIME, ACQ_INS_ID_CD, FWD_INS_ID_CD, OBJ_MCHT_CD, OBJ_TERM_ID, OBJ_MCHT_TP, OBJ_MCHT_NAME, OBJ_MCHT_NAME_OTH, OBJ_ACQ_INS_ID_CD, OBJ_SPEC_FEE_CD, REC_OPR_ID, REC_UPD_OPR, REC_CRT_TS, REC_UPD_TS, AREA_CD) VALUES " \
            "(10, '%s', '*', '* ', '*', '*', '*   ', '*', '*', '*', '*  ', '*', '%s', '*       ', '%s', '*       ', '%s', '* ', '*', '%s', '%s', '%s', '%s','%s', '%s', '*  ', '*', '*         ', sysdate, sysdate, '%s');\n" % \
            (bin_id, high_amt, start_time, end_time, obj_mcht_cd, obj_term_id,mcc_cd, mcht_name, mcht_name, acq_ins_id, area_cd)
        fin.write(sql.encode('gb18030'))

        sql = "INSERT INTO TBL_SHC_MCHT_TERM_MAP (MAP_IDX, BIN_ID, BANK_ID, CARD_TP, CARD_BIN, MCC_TP, MCHT_GP, SRC_MCHT_CD, SRC_TERM_ID, SRC_MCHT_TP, SRC_SPEC_FEE_CD, LOW_AMT, HIGH_AMT, START_DATE, START_TIME, END_DATE, END_TIME, ACQ_INS_ID_CD, FWD_INS_ID_CD, OBJ_MCHT_CD, OBJ_TERM_ID, OBJ_MCHT_TP, OBJ_MCHT_NAME, OBJ_MCHT_NAME_OTH, OBJ_ACQ_INS_ID_CD, OBJ_SPEC_FEE_CD, REC_OPR_ID, REC_UPD_OPR, REC_CRT_TS, REC_UPD_TS, AREA_CD) VALUES " \
            "(20, '%s', '*', '* ', '*', '*', '*   ', '*', '*', '*', '*  ', '*', '%s', '*       ', '%s', '*       ', '%s', '* ', '*', '%s', '%s', '%s', '%s','%s', '%s', '*  ', '*', '*         ', sysdate, sysdate, '%s');\n" % \
            (bin_id, high_amt, start_time, end_time, obj_mcht_cd, obj_term_id, mcc_cd, mcht_name, mcht_name, acq_ins_id, area_cd[0:2])
        fin.write(sql.encode('gb18030'))

        sql = "INSERT INTO TBL_SHC_MCHT_TERM_MAP (MAP_IDX, BIN_ID, BANK_ID, CARD_TP, CARD_BIN, MCC_TP, MCHT_GP, SRC_MCHT_CD, SRC_TERM_ID, SRC_MCHT_TP, SRC_SPEC_FEE_CD, LOW_AMT, HIGH_AMT, START_DATE, START_TIME, END_DATE, END_TIME, ACQ_INS_ID_CD, FWD_INS_ID_CD, OBJ_MCHT_CD, OBJ_TERM_ID, OBJ_MCHT_TP, OBJ_MCHT_NAME,OBJ_MCHT_NAME_OTH, OBJ_ACQ_INS_ID_CD, OBJ_SPEC_FEE_CD, REC_OPR_ID, REC_UPD_OPR, REC_CRT_TS, REC_UPD_TS, AREA_CD) VALUES " \
            "(30, '%s', '*', '* ', '*', '*', '*   ', '*', '*', '*', '*  ', '*', '%s', '*       ', '%s', '*       ', '%s', '* ', '*', '%s', '%s', '%s', '%s','%s', '%s', '*  ', '*', '*         ', sysdate, sysdate, '*');\n" % \
            (bin_id, high_amt, start_time, end_time, obj_mcht_cd, obj_term_id, mcc_cd, mcht_name, mcht_name, acq_ins_id)
        fin.write(sql.encode('gb18030'))


def main():
    fileName = "E:\\需求\\杉德系统数据变更\\商户池\\20190520第十批标扣/第十批标扣 - 副本.xls"
    workbook = xlrd.open_workbook(fileName)
    sh = workbook.sheet_by_index(0)
    nrows = sh.nrows

    i = 1
    while i <  nrows:
        mcht_cd = str(sh.cell(i, 0).value).strip()
        term_id = str(sh.cell(i, 1).value).strip()
        mcht_name = str(sh.cell(i, 5).value).strip()
        mcc_cd = str(sh.cell(i, 7).value)[:4]
        area_cd = str(sh.cell(i, 8).value)[:4]
        acq_ins_id = '4827' + area_cd
        mcht_name = mcht_name.replace("'","‘")
        if len(mcht_name.encode('gb18030')) > 40:
            print("%s,%s" % (mcht_cd, mcht_name))
            i = i + 1
            continue
        genSql(mcc_cd, mcht_cd, term_id, mcht_name, area_cd, acq_ins_id)
        i = i + 1
    fin.close()



if __name__ == '__main__':
    main()