#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
#推送信息

import cx_Oracle
import sys
import os
from hashlib import sha1


def getCheckSum(appSecret, nonce, curTime):
    signdata = appSecret + nonce + curTime
    s1 = sha1()
    s1.update(signdata.encode("gb18030"))
    return s1.hexdigest()

appSecret = 'go9dnk49bkd9jd9vmel1kglw0803mgq3'
nonce = '4tgggergigwow323t23t'
curTime = '1443592222'

print(getCheckSum(appSecret, nonce, curTime))


