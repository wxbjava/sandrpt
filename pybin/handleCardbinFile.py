#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#卡bin比对处理

import xlrd
import re
import cx_Oracle
import os

#acqs序号
idx = 5002


def getInsIdCd(str):
    p = re.compile(r'[(](.*?)[)]', re.S)
    ilen = len(re.findall(p, str))
    return re.findall(p, str)[ilen - 1]

def getFirstLine(str):
    return(str.split()[0])


def checkCardBinExist(db, bin_sta_no):
    cursor = db.cursor()
    sql = "select count(*) from TBL_SHC_CARD_BIN_SWT where CARD_BIN_STA = '%s'" % bin_sta_no
    cursor.execute(sql)
    x = cursor.fetchone()
    cursor.close()
    if x[0] > 0:
        return True
    else:
        return  False


def genPospSql(insIdCd, cardDis, acct1_tnum, acc1_offset, acc1_len,
              bin_tnum, bin_offset, bin_len, bin_sta_no, bin_end_no, card_type):
    if card_type == "贷记卡":
        cardTp = "00"
    elif card_type == "准贷记卡":
        cardTp = "02"
    else:
        cardTp = "01"

    sql = "INSERT INTO TBL_BANK_BIN_INF " \
          "(INS_ID_CD, ACC1_OFFSET, ACC1_LEN, ACC1_TNUM, ACC2_OFFSET, ACC2_LEN, ACC2_TNUM, " \
          "BIN_OFFSET, BIN_LEN, BIN_STA_NO, BIN_END_NO, BIN_TNUM, CARD_TP, CARD_DIS, REC_OPR_ID, " \
          "REC_UPD_OPR, REC_CRT_TS, REC_UPD_TS) VALUES " \
          "('%s', '%s', '%s', '%s', null, null, null, " \
          "'%s', '%s', '%s', '%s', " \
          "'%s', '%s', '%s', null, " \
          "'backsys', sysdate, sysdate);" % (
        insIdCd, acc1_offset, acc1_len, acct1_tnum,
        bin_offset, bin_len, bin_sta_no, bin_end_no, bin_tnum, cardTp, cardDis
    )
    print(sql)


def genAcqSql(insIdCd, cardDis, acct1_tnum, acc1_offset, acc1_len,
              bin_tnum, bin_offset, bin_len, bin_sta_no, bin_end_no, card_type):
    global idx
    idx = idx + 1
    if card_type == "贷记卡":
        cardTp = "00"
    elif card_type == "准贷记卡":
        cardTp = "02"
    else:
        cardTp = "01"

    sql = "INSERT INTO TBL_BANK_BIN_INF " \
          "(IND, INS_ID_CD, ACC1_OFFSET, ACC1_LEN, ACC1_TNUM, ACC2_OFFSET, ACC2_LEN, ACC2_TNUM, " \
          "BIN_OFFSET, BIN_LEN, BIN_STA_NO, BIN_END_NO, BIN_TNUM, CARD_TP, CARD_DIS, REC_OPR_ID, " \
          "REC_UPD_OPR, REC_CRT_TS, REC_UPD_TS) VALUES " \
          "(%d, '%s', '%s', '%s', '%s', null, null, null, " \
          "'%s', '%s', '%s', '%s', " \
          "'%s', '%s', '%s', null, " \
          "'backsys', sysdate, sysdate);" % (
        idx, insIdCd, acc1_offset, acc1_len, acct1_tnum,
        bin_offset, bin_len, bin_sta_no, bin_end_no, bin_tnum, cardTp, cardDis
    )
    print(sql)

def genSwtSql(insIdCd, cardDis, acct1_tnum, acc1_offset, acc1_len,
              bin_tnum, bin_offset, bin_len, bin_sta_no, bin_end_no, card_type):
    global idx
    idx = idx + 1
    if card_type == "贷记卡":
        cardTp = "00"
    elif card_type == "准贷记卡":
        cardTp = "00"
    else:
        cardTp = "01"

    sql = "INSERT INTO TBL_SHC_CARD_BIN (BIN_ID, " \
          "CARD_BIN_LEN, CARD_BIN_STA, CARD_BIN_END, BANK_ID, CARD_TP, FLAGS, " \
          "REC_OPR_ID, REC_UPD_OPR, REC_CRT_TS, REC_UPD_TS) VALUES " \
          "('%s', '%s', '%s', '%s', '%s', '%s', 0, 'a', 'a         ', " \
          "sysdate, sysdate);" % (acc1_len, bin_len, bin_sta_no, bin_end_no, insIdCd, cardTp)
    print(sql)

def main():
    db = cx_Oracle.connect('%s/%s@%s' % (os.environ['DBUSER'], os.environ['DBPWD'], os.environ['TNSNAME']),
                           encoding='gb18030')
    fileName = "../rpt/2018年12月25日版卡表.xls"
    workbook = xlrd.open_workbook(fileName)
    sh = workbook.sheet_by_index(0)
    nrows = sh.nrows
    ncols = sh.ncols
    print(nrows)
    print(ncols)
    i = 4
    while i <  nrows:
        insIdCd = getInsIdCd(sh.cell(i, 0).value)
        cardDis = sh.cell(i, 1).value
        acct1_tnum = getFirstLine(sh.cell(i, 4).value)
        acc1_offset = getFirstLine(sh.cell(i, 5).value)
        acc1_len = getFirstLine(sh.cell(i, 8).value)
        bin_tnum = getFirstLine(sh.cell(i, 10).value)
        bin_offset = getFirstLine(sh.cell(i, 11).value)
        bin_len = getFirstLine(sh.cell(i, 12).value)
        bin_sta_no = getFirstLine(sh.cell(i, 13).value)
        bin_end_no = bin_sta_no
        card_type = getFirstLine(sh.cell(i, 15).value)
        i = i + 1

        #筛选
        if insIdCd == "00010033":
            #银联国际
            continue
        if insIdCd[4:6] >= "01" and insIdCd[4:6] <= "09":
            continue
        res = checkCardBinExist(db, bin_sta_no)
        if res == True:
            continue

        genSwtSql(insIdCd, cardDis, acct1_tnum, acc1_offset, acc1_len,
                  bin_tnum, bin_offset, bin_len, bin_sta_no, bin_end_no, card_type)


if __name__ == '__main__':
    main()