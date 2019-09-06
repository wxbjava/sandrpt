#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-

from hashlib import sha1
from utl.common import *
import random
import string
import requests

class NeteIm:
    def __init__(self, app_secret, app_key, from_user, proxy):
        self.appSecret = app_secret
        self.appKey = app_key
        self.fromUser = from_user
        self.proxy = proxy

    def __get_check_sum(self, nonce, curTime):
        signdata = self.appSecret + nonce + curTime
        s1 = sha1()
        s1.update(signdata.encode("utf-8"))
        return s1.hexdigest()

    def __get_header(self):
        curTime = getCurrUtcStr()
        nonce = ''.join(random.sample(string.ascii_letters + string.digits, 20))
        checksum = self.__get_check_sum(nonce, curTime)
        Content_Type = "application/x-www-form-urlencoded;charset=utf-8"
        header = {'Content-Type': Content_Type, 'AppKey': self.appKey, 'Nonce': str(nonce), 'CurTime': curTime,
                  'CheckSum': checksum}
        return header

    def __get_body(self, to_user, msg):
        data = 'from=%s&ope=0&to=%s&type=0&body={"msg":"%s"}' % (self.fromUser, to_user, msg)
        return data.encode('utf-8')

    def __get_batch_body(self, to_accts, msg):
        toacctlist = ('%s' % to_accts).replace('\'', '"')
        data = 'fromAccid=10000&toAccids=%s&type=0&body={"msg":"%s"}' % (toacctlist, msg)
        return data.encode('utf-8')


    def send_msg(self, to_user, msg):
        url = 'https://api.netease.im/nimserver/msg/sendMsg.action'
        result = []
        try:
            postdata = self.__get_body(to_user, msg)
            result.append(postdata.decode('utf-8'))
            head = self.__get_header()
            response = requests.post(url, data=postdata, headers=head, proxies=self.proxy)
            result.append(response.json())
        except Exception as e:
            result.append(e)
        finally:
            return result

    def send_bat_msg(self, to_user, msg):
        url = 'https://api.netease.im/nimserver/msg/sendBatchMsg.action'
        result = []
        try:
            postdata = self.__get_batch_body(to_user, msg)
            result.append(postdata.decode('utf-8'))
            head = self.__get_header()
            response = requests.post(url, data=postdata, headers=head, proxies=self.proxy)
            result.append(response.json())
        except Exception as e:
            result.append(e)
        finally:
            return result




