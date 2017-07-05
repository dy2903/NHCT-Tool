# -*- coding: UTF-8 -*-  
# from src.package.MySQLdbHelper import  database
from mysqlTools.MySQLDBTools import  database

db=database()
# sql = "SELECT * from listpricetable where BOM = %s ";
sql = "SELECT *  from listpricetable where BOM = %s or typeID = %s";
values = db.fetch_all(sql ,'0150A0X6','NS-SecCenter A2000+BS-1');
for i in values:
    print(i);

db.close();
