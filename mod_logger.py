#encoding:utf-8
import logging.config
from logging.handlers import RotatingFileHandler
import configparser
import os.path
import mod_config


log="log" #日志信息
format       = mod_config.getConfig(log, "format").replace('@', '%')
level         = int(mod_config.getConfig(log, "level"))
backupcount  = int(mod_config.getConfig(log, "backupcount"))
maxbytes     = int(mod_config.getConfig(log, "maxbytes"))

#日志设置
def logger(logpath):
    logger = logging.getLogger(logpath)
    Rthandler = RotatingFileHandler(logpath, maxBytes=maxbytes, backupCount=backupcount)
    #这里来设置日志的级别
    #CRITICAl    50
    #ERROR    40
    #WARNING    30
    #INFO    20
    #DEBUG    10
    #NOSET    0
    #写入日志时，小于指定级别的信息将被忽略
    logger.setLevel(level)
    formatter = logging.Formatter(format)
    Rthandler.setFormatter(formatter)
    logger.addHandler(Rthandler)
    return logger

if __name__ == "__main__":
    print ("mod_logger");
 
LOGPATH = "log.log";
log = logger(LOGPATH);
print(log);
