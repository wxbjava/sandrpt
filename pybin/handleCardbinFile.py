#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#��bin�ȶԴ���

import xlrd
import re


def getInsIdCd(str):
    p = re.compile(r'[(](.*?)[)]', re.S)
    ilen = len(re.findall(p, str))
    return re.findall(p, str)[ilen - 1]

def getFirstLine(str):
    return(str.split()[0])




def main():
    fileName = "../rpt/2018��12��25�հ濨��.xls"
    workbook = xlrd.open_workbook(fileName)
    sh = workbook.sheet_by_index(0)
    nrows = sh.nrows
    ncols = sh.ncols
    print(nrows)
    print(ncols)
    i = 1300
    print(sh.cell(i, 4).value)
    print(getInsIdCd(sh.cell(i, 0).value))
    getFirstLine(sh.cell(i, 4).value)
    i = 4
    while i <  nrows:
        insIdCd = getInsIdCd(sh.cell(i, 0).value)
        cardDis = sh.cell(i, 1).value
        acct1_tnum = getFirstLine(sh.cell(i, 4).value)
        bin_offset = getFirstLine(sh.cell(i, 5).value)
        acc1_len = getFirstLine(sh.cell(i, 8).value)
        bin_tnum = getFirstLine(sh.cell(i, 10).value)
        bin_offset = getFirstLine(sh.cell(i, 11).value)
        bin_len = getFirstLine(sh.cell(i, 12).value)




if __name__ == '__main__':
    main()