#encoding:utf-8
#name:mod_config.py

import configparser
import os

#获取config配置文件
''' 使用方法
dbname = mod_config.getConfig("database", "dbname")
'''
def getConfig(section, key):
    config = configparser.ConfigParser()
    path = os.path.split(os.path.realpath(__file__))[0] + '/configure.conf'
    config.read(path)
    return config.get(section, key)

#其中 os.path.split(os.path.realpath(__file__))[0] 得到的是当前文件模块的目录

# LOGPATH = getConfig('path', 'logpath') + 'database.log';
# print("LOGPATH is ", LOGPATH);
# DB = "database";
# DBNAME = getConfig(DB, 'dbname')
# DBHOST = getConfig(DB, 'dbhost')
# DBUSER = getConfig(DB, 'dbuser')
# DBPWD = getConfig(DB, 'dbpassword')
# DBCHARSET = getConfig(DB, 'dbcharset')
# DBPORT = getConfig(DB, "dbport")

# print(DBNAME);
# print(DBHOST)
# print(DBUSER)
# print(DBPWD )
# print(DBCHARSET)
# print(DBPORT)








