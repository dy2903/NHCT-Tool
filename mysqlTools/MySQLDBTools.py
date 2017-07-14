#!/usr/bin/env python3 
# -*- coding:UTF-8-*-

import mysql.connector
import mod_config
import  mod_logger

# 读取数据库配置
# 注意把password设为你的root口令:
DB = "database"    #数据库配置
LOGPATH = mod_config.getConfig('path', 'logpath') + 'log.log'
DBNAME = mod_config.getConfig(DB, 'dbname')
DBHOST = mod_config.getConfig(DB, 'dbhost')
DBUSER = mod_config.getConfig(DB, 'dbuser')
DBPWD = mod_config.getConfig(DB, 'dbpassword')
DBCHARSET = mod_config.getConfig(DB, 'dbcharset')
DBPORT = mod_config.getConfig(DB, "dbport")

# 初始化日志类
logger = mod_logger.logger(LOGPATH)

# 数据库操作类
class database:
    def __init__(self, dbname=None, dbhost=None):
        self._logger = logger
        #这里的None相当于其它语言的NULL
        if dbname is None:
            self._dbname = DBNAME
        else:
            self._dbname = dbname
        if dbhost is None:
            self._dbhost = DBHOST
        else:
            self._dbhost = dbhost
            
        self._dbuser = DBUSER
        self._dbpassword = DBPWD
        self._dbcharset = DBCHARSET
        self._dbport = int(DBPORT)
        # 调用connectMySQL()函数连接MySQL
        self._conn = self.connectMySQL()
        # 如果连接建立，则获取游标
        if(self._conn):
            self._cursor = self._conn.cursor()

    #连接
    def connectMySQL(self):
        conn = False
        try:
            sqlConfig = {'host':self._dbhost,
                        'port':self._dbport,
                        'user':self._dbuser,
                        'passwd':self._dbpassword,
                        'db':self._dbname,
                        'charset':self._dbcharset
            };        
            conn = mysql.connector.connect(**sqlConfig);
        except Exception as data:
            self._logger.error("connect database failed：%s" % data)
            conn = False
        return conn
 #查询
 # sql：SQL语句
 # *arg：占位符
    def fetch_all(self, sql, *arg):
        res = ''
        if(self._conn):
            try:
                # self._cursor.execute(sql)
                self._cursor.execute(sql,arg)
                res = self._cursor.fetchall()
            except Exception as  data:
                res = False
                self._logger.warn("query database exception, %s" % data)
        return res
    #关闭连接
    def close(self):
        if(self._conn):
            try:
                if(type(self._cursor)=='object'):
                    self._cursor.close()
                if(type(self._conn)=='object'):
                    self._conn.close()
            except Exception as  data:
                self._logger.warn("close database exception, %s,%s,%s" % (data, type(self._cursor), type(self._conn)))
        
