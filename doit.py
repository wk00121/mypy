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
        # 各发卡机构 借记、磁条
        sql = 'SELECT r.`bankcd`, b.banknm, r.`cardtp`, r.`delNum`, r.`delAm` FROM bank b RIGHT JOIN ' \
              '(SELECT SUBSTRING(t.codebank,1,4) AS "bankcd", t.cardRead AS "cardtp", ' \
              'SUM(t.dealNum)/10000 AS "delNum", SUM(t.dealAmount)/1000000 AS "delAm" FROM jourdetail201707 t ' \
              'WHERE t.cardRead IN ("1") ' \
              'AND t.cardType in ("01") ' \
              'GROUP BY SUBSTRING(t.codebank,1,4), t.cardRead ' \
              'ORDER BY SUBSTRING(t.codebank,1,4), t.cardRead) r ' \
              'ON b.`code` = r.`bankcd` ' \
              'GROUP BY b.banknm, r.`cardtp`,b.rank2 ' \
              'ORDER BY b.rank2'
        cursor.execute(sql);
        data = cursor.fetchall()
        ws.write(0, 0, 'bankcd')  # 在1行1列写入发卡行编码
        ws.write(0, 1, 'banknm')  # 在1行2列写入发卡行名称
        ws.write(0, 2, 'cardtp')  # 在1行3列写入卡介质
        ws.write(0, 3, 'delNum')  # 在1行3列写入交易量
        ws.write(0, 4, 'delAm')  # 在1行3列写入交易金额
        for row in data:
            bankcd = row['bankcd']
            banknm = row['banknm']
            cardtp = row['cardtp']
            delNum = row['delNum']
            delAm = row['delAm']
            if banknm == None:
                continue
            i += 1
            ws.write(i, 0, bankcd)  # 在2行1列写入age
            ws.write(i, 1, banknm)  # 在2行2列写入name
            ws.write(i, 2, cardtp)  # 在2行3列写入id
            ws.write(i, 3, delNum)  # 在2行3列写入id
            ws.write(i, 4, delAm)  # 在2行3列写入id
            print str(bankcd) + " " + banknm + " " + str(cardtp) + " " + str(delNum) + " " + str(delAm)
            # 没有设置默认自动提交，需要主动提交，以保存所执行的语句
            # connection.commit()
            ###############################################################################################################2
        i = 0
        data = None
        sql1 = 'SELECT r.`bankcd`, b.banknm, r.`cardtp`, r.`delNum`, r.`delAm` FROM bank b RIGHT JOIN ' \
               '(SELECT SUBSTRING(t.codebank,1,4) AS "bankcd", t.cardRead AS "cardtp", ' \
               'SUM(t.dealNum)/10000 AS "delNum", SUM(t.dealAmount)/1000000 AS "delAm" FROM jourdetail201707 t ' \
               'WHERE t.cardRead IN ("1") ' \
               'AND t.cardType in ("02","03") ' \
               'GROUP BY SUBSTRING(t.codebank,1,4), t.cardRead ' \
               'ORDER BY SUBSTRING(t.codebank,1,4), t.cardRead) r ' \
               'ON b.`code` = r.`bankcd` ' \
               'GROUP BY b.banknm, r.`cardtp`,b.rank2 ' \
               'ORDER BY b.rank2'
        cursor.execute(sql1);
        data = cursor.fetchall()
        ws.write(0, 6, 'bankcd')  # 在1行1列写入发卡行编码
        ws.write(0, 7, 'banknm')  # 在1行2列写入发卡行名称
        ws.write(0, 8, 'cardtp')  # 在1行3列写入卡介质
        ws.write(0, 9, 'delNum')  # 在1行3列写入交易量
        ws.write(0, 10, 'delAm')  # 在1行3列写入交易金额
        for row in data:
            bankcd = row['bankcd']
            banknm = row['banknm']
            cardtp = row['cardtp']
            delNum = row['delNum']
            delAm = row['delAm']
            if banknm == None:
                continue
            i += 1
            ws.write(i, 6, bankcd)  # 在2行1列写入age
            ws.write(i, 7, banknm)  # 在2行2列写入name
            ws.write(i, 8, cardtp)  # 在2行3列写入id
            ws.write(i, 9, delNum)  # 在2行3列写入id
            ws.write(i, 10, delAm)  # 在2行3列写入id
            print str(bankcd) + " " + banknm + " " + str(cardtp) + " " + str(delNum) + " " + str(delAm)

        # connection.commit()
        ###############################################################################################################4
        i = 0
        data = None
        sql1 = 'SELECT r.`bankcd`, b.banknm, r.`cardtp`, r.`delNum`, r.`delAm` FROM bank b RIGHT JOIN ' \
               '(SELECT SUBSTRING(t.codebank,1,4) AS "bankcd", t.cardRead AS "cardtp", ' \
               'SUM(t.dealNum)/10000 AS "delNum", SUM(t.dealAmount)/1000000 AS "delAm" FROM jourdetail201707 t ' \
               'WHERE t.cardRead IN ("3") ' \
               'AND t.cardType in ("01") ' \
               'GROUP BY SUBSTRING(t.codebank,1,4), t.cardRead ' \
               'ORDER BY SUBSTRING(t.codebank,1,4), t.cardRead) r ' \
               'ON b.`code` = r.`bankcd` ' \
               'GROUP BY b.banknm, r.`cardtp`,b.rank2 ' \
               'ORDER BY b.rank2'
        cursor.execute(sql1);
        data = cursor.fetchall()
        ws.write(0, 12, 'bankcd')  # 在1行1列写入发卡行编码
        ws.write(0, 13, 'banknm')  # 在1行2列写入发卡行名称
        ws.write(0, 14, 'cardtp')  # 在1行3列写入卡介质
        ws.write(0, 15, 'delNum')  # 在1行3列写入交易量
        ws.write(0, 16, 'delAm')  # 在1行3列写入交易金额
        for row in data:
            bankcd = row['bankcd']
            banknm = row['banknm']
            cardtp = row['cardtp']
            delNum = row['delNum']
            delAm = row['delAm']
            if banknm == None:
                continue
            i += 1
            ws.write(i, 12, bankcd)  # 在2行1列写入age
            ws.write(i, 13, banknm)  # 在2行2列写入name
            ws.write(i, 14, cardtp)  # 在2行3列写入id
            ws.write(i, 15, delNum)  # 在2行3列写入id
            ws.write(i, 16, delAm)  # 在2行3列写入id
            print str(bankcd) + " " + banknm + " " + str(cardtp) + " " + str(delNum) + " " + str(delAm)
            ###############################################################################################################4
        i = 0
        data = None
        sql1 = 'SELECT r.`bankcd`, b.banknm, r.`cardtp`, r.`delNum`, r.`delAm` FROM bank b RIGHT JOIN ' \
               '(SELECT SUBSTRING(t.codebank,1,4) AS "bankcd", t.cardRead AS "cardtp", ' \
               'SUM(t.dealNum)/10000 AS "delNum", SUM(t.dealAmount)/1000000 AS "delAm" FROM jourdetail201707 t ' \
               'WHERE t.cardRead IN ("3") ' \
               'AND t.cardType in ("02","03") ' \
               'GROUP BY SUBSTRING(t.codebank,1,4), t.cardRead ' \
               'ORDER BY SUBSTRING(t.codebank,1,4), t.cardRead) r ' \
               'ON b.`code` = r.`bankcd` ' \
               'GROUP BY b.banknm, r.`cardtp`,b.rank2 ' \
               'ORDER BY b.rank2'
        cursor.execute(sql1);
        data = cursor.fetchall()
        ws.write(0, 18, 'bankcd')  # 在1行1列写入发卡行编码
        ws.write(0, 19, 'banknm')  # 在1行2列写入发卡行名称
        ws.write(0, 20, 'cardtp')  # 在1行3列写入卡介质
        ws.write(0, 21, 'delNum')  # 在1行3列写入交易量
        ws.write(0, 22, 'delAm')  # 在1行3列写入交易金额
        for row in data:
            bankcd = row['bankcd']
            banknm = row['banknm']
            cardtp = row['cardtp']
            delNum = row['delNum']
            delAm = row['delAm']
            if banknm == None:
                continue
            i += 1
            ws.write(i, 18, bankcd)  # 在2行1列写入age
            ws.write(i, 19, banknm)  # 在2行2列写入name
            ws.write(i, 20, cardtp)  # 在2行3列写入id
            ws.write(i, 21, delNum)  # 在2行3列写入id
            ws.write(i, 22, delAm)  # 在2行3列写入id
            print str(bankcd) + " " + banknm + " " + str(cardtp) + " " + str(delNum) + " " + str(delAm)
            # 没有设置默认自动提交，需要主动提交，以保存所执行的语句
    connection.commit()
    # 各发卡机构 贷记、磁条


finally:
    connection.close();
w.save('yinhang.xls')  # 保存