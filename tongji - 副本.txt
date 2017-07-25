# -*- coding: utf-8 -*-1

import pymysql.cursors
from pyExcelerator import *
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

# 函数:执行地市sql
def toDodishi(a,b,c):
    # 各地市sql语句

    sql = 'SELECT b.bank , t.cardRead , SUM(t.dealNum) as "bishu" , ' \
          'SUM(t.dealAmount)/1000000 as "jine"  FROM jourdetail201706 t INNER JOIN bank b ' \
          'ON SUBSTRING(t.codebank,1,4) = b.`code` WHERE SUBSTRING(t.acceptInstitution, 5, 3) IN ("364")' \
          ' AND t.cardRead IN ("1", "3") AND t.cardType IN ("02","03")  GROUP BY' \
          ' b.bank, t.cardRead ORDER BY b.bank, t.cardRead'

    return a


i = 0  # 声明循环计数器
config = {  # 连接服务器
    'host': '127.0.0.1',
    'port': 3306,
    'user': 'root',
    'password': '',
    'db': 'yinhang',
    'charset': 'utf8mb4',
    'cursorclass': pymysql.cursors.DictCursor,
}
# Connect to the database
connection = pymysql.connect(**config)

# 执行sql语句
w = Workbook()  # 创建一个工作簿
ws = w.add_sheet('aaa')  # 创建一个工作表
try:

    with connection.cursor() as cursor:
        # 执行sql语句，插入记录
        #各地市sql语句
        sql = 'SELECT b.bank , t.cardRead , SUM(t.dealNum) as "bishu" , ' \
              'SUM(t.dealAmount)/1000000 as "jine"  FROM jourdetail201706 t INNER JOIN bank b ' \
              'ON SUBSTRING(t.codebank,1,4) = b.`code` WHERE SUBSTRING(t.acceptInstitution, 5, 3) IN ("364")' \
              ' AND t.cardRead IN ("1", "3") AND t.cardType IN ("02","03")  GROUP BY' \
              ' b.bank, t.cardRead ORDER BY b.bank, t.cardRead'

        cursor.execute(sql);
        data = cursor.fetchall()
        ws.write(0, 0, 'bank')  # 在1行1列写入age
        ws.write(0, 1, 'cardread')  # 在1行2列写入name
        ws.write(0, 2, 'delNum')  # 在1行3列写入id
        ws.write(0, 3, 'delAm')  # 在1行3列写入id
        for row in data:
            bank = row['bank']
            cardRead = row['cardRead']
            dealNum = row['bishu']
            dealAm = row['jine']
            i += 1
            ws.write(i, 0, bank)  # 在2行1列写入age
            ws.write(i, 1, cardRead)  # 在2行2列写入name
            ws.write(i, 2, dealNum)  # 在2行3列写入id
            ws.write(i, 3, dealAm)  # 在2行3列写入id
            print str(bank) + " " + cardRead + " " + str(dealNum)
    # 没有设置默认自动提交，需要主动提交，以保存所执行的语句
    connection.commit()
finally:
    connection.close();
w.save('mini.xls')  # 保存