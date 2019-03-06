#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-

import xlrd
import cx_Oracle
import os

bin_id = '9892'
db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSERSWT'], os.environ['DBPWDSWT'], os.environ['TNSNAME']),
                           encoding='gb18030')

filename = "E:\\����\\ɼ��ϵͳ���ݱ��\\20190301���������/data20190301.sql"
fin = open(filename, "wb")

def getMccInfo(mcc_cd):
    sql = "select HIGH_AMT, BEGIN_TIME, END_TIME from TBL_SHC_MCC_CFG where MCC_CD ='%s'" % mcc_cd
    cursor = db.cursor()
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    if x is not None:
        return x[0],x[1],x[2]
    else:
        print("δ�鵽")
        return '*','*','*'

def genSql(mcc_cd, obj_mcht_cd, obj_term_id, mcht_name, area_cd, acq_ins_id):
    sql = "INSERT INTO TBL_SHC_MCHT_TERM_MAP (MAP_IDX, BIN_ID, BANK_ID, CARD_TP, CARD_BIN, MCC_TP, MCHT_GP, SRC_MCHT_CD, SRC_TERM_ID, SRC_MCHT_TP, SRC_SPEC_FEE_CD, LOW_AMT, HIGH_AMT, START_DATE, START_TIME, END_DATE, END_TIME, ACQ_INS_ID_CD, FWD_INS_ID_CD, OBJ_MCHT_CD, OBJ_TERM_ID, OBJ_MCHT_TP, OBJ_MCHT_NAME, OBJ_MCHT_NAME_OTH, OBJ_ACQ_INS_ID_CD, OBJ_SPEC_FEE_CD, REC_OPR_ID, REC_UPD_OPR, REC_CRT_TS, REC_UPD_TS, AREA_CD) VALUES " \
          "(5, '%s', '*', '* ', '*', '*', '*', '*', '*', '%s', '*  ', '*', '*', '*', '*', '*', '*', '09970000', '*          ', '%s', '%s', '%s', '%s','%s', '%s', '*  ', '*', '*         ', sysdate, sysdate, '%s');\n" % (bin_id, mcc_cd, obj_mcht_cd, obj_term_id,mcc_cd, mcht_name, mcht_name, acq_ins_id, area_cd)
    fin.write(sql.encode('gb18030'))

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
    fileName = "E:\\����\\ɼ��ϵͳ���ݱ��\\20190301���������/���������(1).xls"
    workbook = xlrd.open_workbook(fileName)
    sh = workbook.sheet_by_index(0)
    nrows = sh.nrows

    i = 3
    while i <  nrows:
        mcht_cd = sh.cell(i, 0).value
        term_id = sh.cell(i, 1).value
        mcht_name = sh.cell(i, 4).value
        mcc_cd = sh.cell(i, 5).value
        area_cd = str(sh.cell(i, 6).value)[:4]
        acq_ins_id = '4827' + area_cd
        if len(mcht_name.encode('gb18030')) > 40:
            print("%s,%s" % (mcht_cd, mcht_name))
            i = i + 1
            continue
        genSql(mcc_cd, mcht_cd, term_id, mcht_name, area_cd, acq_ins_id)
        i = i + 1
    fin.close()



if __name__ == '__main__':
    main()