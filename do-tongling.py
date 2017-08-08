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
ws = w.add_sheet('huaibei')  # 创建一个工作表
try:
    with connection.cursor() as cursor:
        # 执行sql语句，插入记录
        # 各地市 磁条、借记
        sql = 'SELECT b.banknm AS "banknm", t.cardRead AS "cardtp", ' \
              'SUM(t.dealNum) AS "delNum", SUM(t.dealAmount)/1000000 AS "delAm" ' \
              'FROM jourdetail201707 t INNER JOIN bank b ON SUBSTRING(t.codebank,1,4) = b.`code` ' \
              'WHERE SUBSTRING(t.acceptInstitution, 5, 3) IN ("367") ' \
              'AND t.cardRead IN ("1")' \
              'AND t.cardType IN ("01")' \
              ' GROUP BY b.banknm, t.cardRead,b.rank' \
              ' ORDER BY b.rank'

        cursor.execute(sql);
        data = cursor.fetchall()
        ws.write(0, 0, 'banknm')  # 在1行1列写入发卡行编码
        ws.write(0, 1, 'cardtp')  # 在1行2列写入发卡行名称
        ws.write(0, 2, 'delNum')  # 在1行3列写入卡介质
        ws.write(0, 3, 'delAm')  # 在1行3列写入交易量
        for row in data:
            banknm = row['banknm']
            cardtp = row['cardtp']
            delNum = row['delNum']
            delAm = row['delAm']
            if banknm == None:
                continue
            i += 1
            ws.write(i, 0, banknm)  # 在2行1列写入age
            ws.write(i, 1, cardtp)  # 在2行3列写入id
            ws.write(i, 2, delNum)  # 在2行3列写入id
            ws.write(i, 3, delAm)  # 在2行3列写入id
            print str(banknm) + " " + cardtp + " " + str(delNum) + " " + str(delAm) + " "

            ###############################################################################################################2
        i = 0
        data = None
        sql1 ='SELECT b.banknm AS "banknm", t.cardRead AS "cardtp", ' \
              'SUM(t.dealNum) AS "delNum", SUM(t.dealAmount)/1000000 AS "delAm" ' \
              'FROM jourdetail201707 t INNER JOIN bank b ON SUBSTRING(t.codebank,1,4) = b.`code` ' \
              'WHERE SUBSTRING(t.acceptInstitution, 5, 3) IN ("367") ' \
              'AND t.cardRead IN ("1")' \
              'AND t.cardType IN ("02","03")' \
              ' GROUP BY b.banknm, t.cardRead,b.rank' \
              ' ORDER BY b.rank'

        cursor.execute(sql1);
        data = cursor.fetchall()
        ws.write(0, 6, 'banknm')  # 在1行1列写入发卡行编码
        ws.write(0, 7, 'cardtp')  # 在1行2列写入发卡行名称
        ws.write(0, 8, 'delNum')  # 在1行3列写入卡介质
        ws.write(0, 9, 'delAm')  # 在1行3列写入交易量
        for row in data:
            banknm = row['banknm']
            cardtp = row['cardtp']
            delNum = row['delNum']
            delAm = row['delAm']
            if banknm == None:
                continue
            i += 1
            ws.write(i, 6, banknm)  # 在2行1列写入age
            ws.write(i, 7, cardtp)  # 在2行2列写入name
            ws.write(i, 8, delNum)  # 在2行3列写入id
            ws.write(i, 9, delAm)  # 在2行3列写入id
            print str(banknm) + " " + cardtp + " " + str(delNum) + " " + str(delAm) + " "

        ###############################################################################################################2
        i = 0
        data = None
        sql2= 'SELECT b.banknm AS "banknm", t.cardRead AS "cardtp", ' \
              'SUM(t.dealNum) AS "delNum", SUM(t.dealAmount)/1000000 AS "delAm" ' \
              'FROM jourdetail201707 t INNER JOIN bank b ON SUBSTRING(t.codebank,1,4) = b.`code` ' \
              'WHERE SUBSTRING(t.acceptInstitution, 5, 3) IN ("367") ' \
              'AND t.cardRead IN ("3")' \
              'AND t.cardType IN ("01")' \
              ' GROUP BY b.banknm, t.cardRead,b.rank' \
              ' ORDER BY b.rank'

        cursor.execute(sql2);
        data = cursor.fetchall()
        ws.write(0, 11, 'banknm')  # 在1行1列写入发卡行编码
        ws.write(0, 12, 'cardtp')  # 在1行2列写入发卡行名称
        ws.write(0, 13, 'delNum')  # 在1行3列写入卡介质
        ws.write(0, 14, 'delAm')  # 在1行3列写入交易量

        for row in data:
            banknm = row['banknm']
            cardtp = row['cardtp']
            delNum = row['delNum']
            delAm = row['delAm']

            if banknm == None:
                continue
            i += 1
            ws.write(i, 11, banknm)  # 在2行1列写入age
            ws.write(i, 12, cardtp)  # 在2行2列写入name
            ws.write(i, 13, delNum)  # 在2行3列写入id
            ws.write(i, 14, delAm)  # 在2行3列写入id
            print str(banknm) + " " + cardtp + " " + str(delNum) + " " + str(delAm) + " "

        ###############################################################################################################2
        i = 0
        data = None
        sql3 ='SELECT b.banknm AS "banknm", t.cardRead AS "cardtp", ' \
              'SUM(t.dealNum) AS "delNum", SUM(t.dealAmount)/1000000 AS "delAm" ' \
              'FROM jourdetail201707 t INNER JOIN bank b ON SUBSTRING(t.codebank,1,4) = b.`code` ' \
              'WHERE SUBSTRING(t.acceptInstitution, 5, 3) IN ("367") ' \
              'AND t.cardRead IN ("3")' \
              'AND t.cardType IN ("02","03")' \
              ' GROUP BY b.banknm, t.cardRead,b.rank' \
              ' ORDER BY b.rank'
        cursor.execute(sql3);
        data = cursor.fetchall()
        ws.write(0, 16, 'banknm')  # 在1行1列写入发卡行编码
        ws.write(0, 17, 'cardtp')  # 在1行2列写入发卡行名称
        ws.write(0, 18, 'delNum')  # 在1行3列写入卡介质
        ws.write(0, 19, 'delAm')  # 在1行3列写入交易量

        for row in data:
            banknm = row['banknm']
            cardtp = row['cardtp']
            delNum = row['delNum']
            delAm = row['delAm']
            if banknm == None:
                continue
            i += 1
            ws.write(i, 16, banknm)  # 在2行1列写入age
            ws.write(i, 17, cardtp)  # 在2行2列写入name
            ws.write(i, 18, delNum)  # 在2行3列写入id
            ws.write(i, 19, delAm)  # 在2行3列写入id

            print str(banknm) + " " + cardtp + " " + str(delNum) + " " + str(delAm) + " "

            # connection.commit()
    # 没有设置默认自动提交，需要主动提交，以保存所执行的语句
    connection.commit()
finally:
    connection.close();
w.save('dishi-tongling.xls')  # 保存