# coding=utf-8
# 快递100 API
# 作者: 你大爷
# 邮箱: yourdaye@gmail.com
# 介绍 次程序只需输入快递运单编号，就可以返回快递的详细信息
import requests
import urllib
import sys
reload(sys)
sys.setdefaultencoding('utf8')

# 函数: 承运公司名到文本
def GetComName(comCode):
    if comCode=='shentong':
        return '申通快递'
    elif comCode=='zhontong':
        return '中通快递'
    elif comCode=='ems':
        return 'EMS'
    elif comCode=='huitongkuaidi':
        return '汇通快运'
    else:
        return comCode

# 函数: 取状态文本
def GetStateText(num):
    if num==0:
        return '运输中'
    elif num==1:
        return '揽件'
    elif num==2:
        return '疑难'
    elif num==3:
        return '已签收'
    elif num==4:
        return '退回签收'
    elif num==5:
        return '派送中'
    elif num==6:
        return '退回中'

p = {}
p['text'] = input("请输入快递运单编号: ")  #比如: 227728570825
autoComNum = requests.get("http://www.kuaidi100.com/autonumber/autoComNum", params=p)
com = autoComNum.json()


if com['auto'] == []:
    print("这是一个错误的运单编号!")

else:
    print("\n---------------- 承运公司 ------------------\n")
    i=0
    for this in com['auto']:
        i = i + 1
        print( str(i) + ". " + GetComName(this['comCode']) + "\n")


    num = input("承运公司序号: ")
    print("\n---------------- 正在查询, 请稍等... ------------------\n")
    data = {}
    data['type'] = com['auto'][int(num)-1]['comCode']
    data['postid'] =  p['text']
    data['valicode'] = ''
    data['id'] = 1
    data['temp'] = '0.14881871496191512'
    query = requests.get("http://www.kuaidi100.com/query", params=data)
    res = query.json()

    print("\n运单编号 --> " + res['nu'])
    print("\n承运公司 --> " + GetComName(res['com']))
    print("\n当前状态 --> " + GetStateText(int(res['state'])))
    print("\n---------------- 跟踪信息 ------------------\n")
    for this in res['data']:
        print(this['time'] + "\t" + this['context'] + "\n")

input("\n\n感谢使用, 脚本作者:nidaye.   API:快递100")