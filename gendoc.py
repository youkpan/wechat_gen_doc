#!/usr/bin/env python
# -*- coding: utf-8 -*-
import MySQLdb as mdb
import docx
from docx.shared import Inches
import mysql.connector

mydb = mysql.connector.connect(
  host="v.hayoou.com",       # 数据库主机地址
  user="root",    # 数据库用户名
  passwd="yourpassword"   # 数据库密码
)

mycursor = mydb.cursor()
 
mycursor.execute(" ")
myresult = mycursor.fetchall()     # fetchall() 获取所有记录
 
for x in myresult:
  print(x)


#新建文档
doc_new = docx.Document()
document.add_picture('monty-truth.png', width=Inches(1.25))
doc_new.save("./gen/1.docx")

