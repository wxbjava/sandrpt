#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-
import socket
import os


class sndsms:
    def __init__(self):
        self.__ip = os.environ['SMSIP']
        self.__port = int(os.environ['SMSPORT'])

    def setPhone(self, phoneList):
        self.__phlst = phoneList

    def getPhone(self):
        return self.__phlst

    def setMsg(self, msg):
        self.msg = msg

    def __send(self, phone):
        head = "07jmk                           37                  sandqgd   qawsed12  "
        sendStr = head + phone + self.msg
        sendByte = (sendStr.ljust(327 - len(sendStr.encode('gb18030')) + len(sendStr))).encode('gb18030')
        client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        client.connect((self.__ip, self.__port))
        client.send(sendByte)
        data = client.recv(9999)
        client.close()
        return data.decode('gb18030')

    def sndSms(self):
        for phone in self.__phlst:
            self.__send(phone)


