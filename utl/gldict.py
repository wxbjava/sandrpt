#-*- coding:gb18030 -*-

itemDict = {
    '325': '无标识',
    '402': '点点客',
    '543': '北京医院',
    '544': '卡耐丁',
    '553': '银联供应链',
    '638': '福农通',
    '639': '弘付',
    '640': '宁波银嘉',
    '641': '广州银嘉',
    '642': '辽宁直联T0',
    '643': '现场注册',
    '644': '瀚银',
    '742': '杉德多多付',
    '743': '银拓',
    '886': '银联线上',
    '898': '弘付49',
    '899': '瀚银49',
    '900': '广东银嘉49',
    '901': '清算-台州',
    '902': '清算-村镇',
    '904': '弈轩二维码',
    '1129': '恒大集团',
    '1130': '间联弘付',
    '1131': '间联弘付49',
    '1132': '间联广东银嘉',
    '1133': '间联广东银嘉49',
    '1134': '间联瀚银',
    '1135': '间联瀚银49',
    '1136': '间联久付',
    '1137': '间联久付49',
    '1138': '间联久久付',
    '1139': '直联弘付',
    '1156': '传统-00（东北分公司专用）',
    '1157': '北京一鸣(东北分公司专用)',
    '1158': '网络商城（东北分公司专用）',
    '1172': '杉德瀚银',
    '1173': '杉德瀚银49',
    '1182': '银联赈灾',
    '1183': '间联冠儒1',
    '1184': '间联冠儒2',
    '1185': '卡说二维码',
    '1187': '中石化',
    '1188': '辽宁直联T0-凯富（东北专用）',
    '1189': 'CBXW',
    '1581': '小微商户',
    '1591': '间联久付减免',
    '1596': '间联广州银嘉',
    '1597': '间联东北优惠',
    '1598': '间联东北减免',
    '1599': '间联东北49',
    '1600': '间联山西优惠',
    '1601': '间联山西减免',
    '1602': '间联山西49',
    '1603': '间联浙江优惠',
    '1604': '间联浙江减免',
    '1605': '间联浙江49',
    '1606': '瀚银久久付',
    '1608': '哆啦云49',
    '1609': '哆啦云标准',
    '1610': '哆啦云商旅',
    '1668': '直联瀚银',
    '1669': '间联弘付优惠',
    '1670': '间联弘付减免',
    '1674': '二维码支付',
    '1675': '间联东北标准',
    '1676': '间联广东',
    '1677': '间联广东优惠',
    '1678': '间联广东49',
    '1679': '间联广东减免',
    '1701': '网络支付',
    '1789': '交行聚合',
    '1790': '瀚银标准',
    '1791': '瀚银优惠',
    '1792': '瀚银减免',
    '1793': '弘付信用卡还款',
    '1794': '弘付标准',
    '1795': '弘付49-线上',
    '1796': '弘付优惠',
    '1797': '弘付减免',
    '1798': '弘付商旅-线上',
    '1799': '弘付商旅-线上1',
    '1800': '弘付标准-线上',
    '1801': '弘付烟草-线上',
    '1802': '间联CBXW标准',
    '1803': '间联CBXW优惠',
    '1804': '间联CBXW49',
    '1805': '间联CBXW减免',
    '1808': '聚财通',
    '1809': '大连出租车',
    '1821': 'SD-线上标准',
    '1822': 'SD-线上49',
    '1823': 'SD-线上商旅',
    '1824': '间联山西标准',
    '1843': '助农取款',
    '1862': '间联闪付',
    '1894': '间联青岛标准',
    '1895': '间联青岛优惠',
    '1896': '间联青岛49',
    '1897': '间联青岛减免',
    '1929': '直联云闪付',
    '2122': '间联商户'
}

#获取项目标识中文名
def getItemName(itemId):
    value = itemId.rstrip()
    try:
        name = itemDict[value]
    except KeyError:
        name = "未知项目(%s)" % value
    return name

#通道id字典
chnlDict = {
    '00000160' : '上海直联通道',
    '00000440' : '交行银企通道',
    '00000730' : '平安垫资代付通道',
    '00000991' : '合肥平安T1代付通道'
}

#获取通道中文名
def getChnlName(chnlId):
    value = chnlId.rstrip()
    try:
        name = chnlDict[value]
    except KeyError:
        name = "未知通道(%s)" % value
    return name