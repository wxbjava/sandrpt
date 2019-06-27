#!/home/acqbat/python36/bin/python3
#-*- coding:gb18030 -*-

import utl.properties as properties
from aliyunsdkpush.request.v20160801 import PushRequest
from aliyunsdkcore import client
from datetime import *
import time

clt = client.AcsClient(properties.accessKeyId,properties.accessKeySecret,properties.regionId)

request = PushRequest.PushRequest()
request.set_AppKey(properties.appKey)
#����Ŀ��: DEVICE:���豸���� ALIAS : ���������� ACCOUNT:���ʺ�����  TAG:����ǩ����; ALL: �㲥����
request.set_Target('ACCOUNT')
#����Target���趨����Target=DEVICE, ���Ӧ��ֵΪ �豸id1,�豸id2. ���ֵʹ�ö��ŷָ�.(�ʺ����豸��һ�����100��������)
request.set_TargetValue("301566")
#�豸���� ANDROID iOS ALL
request.set_DeviceType("ALL")
#��Ϣ���� MESSAGE NOTICE
request.set_PushType("MESSAGE")
#��Ϣ�ı���
request.set_Title("Open Api Push Title")
#��Ϣ������
request.set_Body("{\"msg\":\"������Ϣ����,����id301566\"}")

# iOS����

#iOSӦ��ͼ�����ϽǽǱ�
request.set_iOSBadge(5)
#������Ĭ֪ͨ
request.set_iOSSilentNotification(False)
#iOS֪ͨ����
request.set_iOSMusic("default")
#iOS��֪ͨ��ͨ��APNs���������͵ģ���Ҫ��д��Ӧ�Ļ�����Ϣ��"DEV" : ��ʾ�������� "PRODUCT" : ��ʾ��������
request.set_iOSApnsEnv("PRODUCT")
# ��Ϣ����ʱ�豸�����ߣ������ƶ����͵ķ���˵ĳ�����ͨ����ͨ�������������ͻ���Ϊ֪ͨ��ͨ��ƻ����APNsͨ���ʹ�һ�Ρ�ע�⣺������Ϣת֪ͨ����������������
request.set_iOSRemind(True)
#iOS��Ϣת֪ͨʱʹ�õ�iOS֪ͨ���ݣ�����iOSApnsEnv=PRODUCT && iOSRemindΪtrueʱ��Ч
request.set_iOSRemindBody("iOSRemindBody");
#�Զ����kv�ṹ,��������չ�� ���iOS�豸
request.set_iOSExtParameters("{\"k1\":\"ios\",\"k2\":\"v2\"}")

#android����

#֪ͨ�����ѷ�ʽ "VIBRATE" : �� "SOUND" : ���� "BOTH" : �������� NONE : ����
request.set_AndroidNotifyType("BOTH")
#֪ͨ���Զ�����ʽ1-100
request.set_AndroidNotificationBarType(1)
#���֪ͨ���� "APPLICATION" : ��Ӧ�� "ACTIVITY" : ��AndroidActivity "URL" : ��URL "NONE" : ����ת
request.set_AndroidOpenType("APPLICATION");
#Android�յ����ͺ�򿪶�Ӧ��url,����AndroidOpenType="URL"��Ч
#request.set_AndroidOpenUrl("www.aliyun.com")
#�趨֪ͨ�򿪵�activity������AndroidOpenType="Activity"��Ч
#request.set_AndroidActivity("com.alibaba.push2.demo.XiaoMiPushActivity")
#Android֪ͨ����
request.set_AndroidMusic("default")
#���øò���������С���йܵ�������, �˴�ָ��֪ͨ�������ת��Activity���йܵ�����ǰ��������1. ����С�׸���ͨ����2. StoreOffline������Ϊtrue��
request.set_AndroidXiaoMiActivity("com.alibaba.push2.demo.XiaoMiPushActivity")
request.set_AndroidXiaoMiNotifyTitle("Mi title")
request.set_AndroidXiaoMiNotifyBody("MiActivity Body")
#�趨֪ͨ����չ���ԡ�(ע�� : �ò���Ҫ�� json map �ĸ�ʽ����,������������)
request.set_AndroidExtParameters("{\"k1\":\"android\",\"k2\":\"v2\"}")

#���Ϳ���
#30��֮����, Ҳ�������ó���ָ���̶�ʱ��
pushDate = datetime.utcnow() + timedelta(seconds = +1)
#24Сʱ����ϢʧЧ, �����ٷ���
expireDate = datetime.utcnow() + timedelta(hours = +24)
#ת����ISO8601T���ݸ�ʽ
pushTime = pushDate.strftime("%Y-%m-%dT%XZ")
expireTime = expireDate.strftime("%Y-%m-%dT%XZ")
request.set_PushTime(pushTime)
request.set_ExpireTime(expireTime)
#���ù���ʱ�䣬��λ��Сʱ
#request.set_TimeOut(24)
request.set_StoreOffline(True)

result = clt.do_action_with_exception(request)
print(result)
