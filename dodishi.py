# -*- coding: utf-8 -*-1

import pymysql.cursors
from pyExcelerator import *
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

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
        # 各地市 磁条、借记
        sql = ' SELECT r.`dishicd`,b.`citynm`,r.`delNum`,r.`delAm`,b.rank from citys b right join' \
              '(select SUBSTRING(t.acceptInstitution, 5, 3) AS "dishicd", t.cardRead AS "cardtp", ' \
              'SUM(t.dealNum)/10000 AS "delNum", SUM(t.dealAmount)/1000000 AS "delAm"  ' \
              ' FROM jourdetail201707 t ' \
              'WHERE t.cardType IN ("01") ' \
              'AND t.cardRead IN ("1") ' \
              ' GROUP BY SUBSTRING(t.acceptInstitution, 5, 3), t.cardRead ) r' \
              ' on b.`ctnum` = r.`dishicd`' \
              'GROUP BY b.citynm, b.ctnum,b.rank ORDER BY b.rank'

        cursor.execute(sql);
        data = cursor.fetchall()
        ws.write(0, 0, 'dishicd')  # 在1行1列写入发卡行编码
        ws.write(0, 1, 'citynm')  # 在1行2列写入发卡行名称
        ws.write(0, 2, 'delNum')  # 在1行3列写入卡介质
        ws.write(0, 3, 'delAm')  # 在1行3列写入交易量
        for row in data:
            dishicd = row['dishicd']
            citynm = row['citynm']
            delNum = row['delNum']
            delAm = row['delAm']
            if citynm == None:
                continue
            i += 1
            ws.write(i, 0, dishicd)  # 在2行1列写入age
            ws.write(i, 1, citynm)  # 在2行3列写入id
            ws.write(i, 2, delNum)  # 在2行3列写入id
            ws.write(i, 3, delAm)  # 在2行3列写入id
            print str(dishicd) + " " + citynm + " " + str(delNum) + " " + str(delAm) + " "

            ###############################################################################################################2
        i = 0
        data = None
        sql1 =' SELECT r.`dishicd`,b.`citynm`,r.`delNum`,r.`delAm`,b.rank from citys b right join' \
              '(select SUBSTRING(t.acceptInstitution, 5, 3) AS "dishicd", t.cardRead AS "cardtp", ' \
              'SUM(t.dealNum)/10000 AS "delNum", SUM(t.dealAmount)/1000000 AS "delAm"  ' \
              ' FROM jourdetail201707 t ' \
              'WHERE t.cardType IN ("02","03") ' \
              'AND t.cardRead IN ("1") ' \
              ' GROUP BY SUBSTRING(t.acceptInstitution, 5, 3), t.cardRead ) r' \
              ' on b.`ctnum` = r.`dishicd`' \
              'GROUP BY b.citynm, b.ctnum,b.rank ORDER BY b.rank'

        cursor.execute(sql1);
        data = cursor.fetchall()
        ws.write(0, 6, 'dishicd')  # 在1行1列写入发卡行编码
        ws.write(0, 7, 'citynm')  # 在1行2列写入发卡行名称
        ws.write(0, 8, 'delNum')  # 在1行3列写入卡介质
        ws.write(0, 9, 'delAm')  # 在1行3列写入交易量
        for row in data:
            dishicd = row['dishicd']
            citynm = row['citynm']
            delNum = row['delNum']
            delAm = row['delAm']
            if citynm == None:
                continue
            i += 1
            ws.write(i, 6, dishicd)  # 在2行1列写入age
            ws.write(i, 7, citynm)  # 在2行2列写入name
            ws.write(i, 8, delNum)  # 在2行3列写入id
            ws.write(i, 9, delAm)  # 在2行3列写入id
            print str(dishicd) + " " + citynm + " " + str(delNum) + " " + str(delAm) + " "

        ###############################################################################################################2
        i = 0
        data = None
        sql2= ' SELECT r.`dishicd`,b.`citynm`,r.`delNum`,r.`delAm`,b.rank from citys b right join' \
              '(select SUBSTRING(t.acceptInstitution, 5, 3) AS "dishicd", t.cardRead AS "cardtp", ' \
              'SUM(t.dealNum)/10000 AS "delNum", SUM(t.dealAmount)/1000000 AS "delAm"  ' \
              ' FROM jourdetail201707 t ' \
              'WHERE t.cardType IN ("01") ' \
              'AND t.cardRead IN ("3") ' \
              ' GROUP BY SUBSTRING(t.acceptInstitution, 5, 3), t.cardRead ) r' \
              ' on b.`ctnum` = r.`dishicd`' \
              'GROUP BY b.citynm, b.ctnum,b.rank ORDER BY b.rank'
        cursor.execute(sql2);
        data = cursor.fetchall()
        ws.write(0, 11, 'dishicd')  # 在1行1列写入发卡行编码
        ws.write(0, 12, 'citynm')  # 在1行2列写入发卡行名称
        ws.write(0, 13, 'delNum')  # 在1行3列写入卡介质
        ws.write(0, 14, 'delAm')  # 在1行3列写入交易量

        for row in data:
            dishicd = row['dishicd']
            citynm = row['citynm']
            delNum = row['delNum']
            delAm = row['delAm']

            if citynm == None:
                continue
            i += 1
            ws.write(i, 11, dishicd)  # 在2行1列写入age
            ws.write(i, 12, citynm)  # 在2行2列写入name
            ws.write(i, 13, delNum)  # 在2行3列写入id
            ws.write(i, 14, delAm)  # 在2行3列写入id
            print str(dishicd) + " " + citynm + " " + str(delNum) + " " + str(delAm) + " "

        ###############################################################################################################2
        i = 0
        data = None
        sql3 = ' SELECT r.`dishicd`,b.`citynm`,r.`delNum`,r.`delAm`,b.rank from citys b right join' \
              '(select SUBSTRING(t.acceptInstitution, 5, 3) AS "dishicd", t.cardRead AS "cardtp", ' \
              'SUM(t.dealNum)/10000 AS "delNum", SUM(t.dealAmount)/1000000 AS "delAm"  ' \
              ' FROM jourdetail201707 t ' \
              'WHERE t.cardType IN ("02","03") ' \
              'AND t.cardRead IN ("3") ' \
              ' GROUP BY SUBSTRING(t.acceptInstitution, 5, 3), t.cardRead ) r' \
              ' on b.`ctnum` = r.`dishicd`' \
              'GROUP BY b.citynm, b.ctnum,b.rank ORDER BY b.rank'
        cursor.execute(sql3);
        data = cursor.fetchall()
        ws.write(0, 16, 'dishicd')  # 在1行1列写入发卡行编码
        ws.write(0, 17, 'citynm')  # 在1行2列写入发卡行名称
        ws.write(0, 18, 'delNum')  # 在1行3列写入卡介质
        ws.write(0, 19, 'delAm')  # 在1行3列写入交易量

        for row in data:
            dishicd = row['dishicd']
            citynm = row['citynm']
            delNum = row['delNum']
            delAm = row['delAm']
            if citynm == None:
                continue
            i += 1
            ws.write(i, 16, dishicd)  # 在2行1列写入age
            ws.write(i, 17, citynm)  # 在2行2列写入name
            ws.write(i, 18, delNum)  # 在2行3列写入id
            ws.write(i, 19, delAm)  # 在2行3列写入id

            print str(dishicd) + " " + citynm + " " + str(delNum) + " " + str(delAm) + " "

            # connection.commit()
    # 没有设置默认自动提交，需要主动提交，以保存所执行的语句
    connection.commit()
finally:
    connection.close();
w.save('dishi.xls')  # 保存