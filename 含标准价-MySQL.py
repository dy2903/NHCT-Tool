#/usr/bin/env python3
# -*- coding:UTF-8-*-
from quotationTools.excelTools  import excelTools
import os.path
import re

import mod_config
import mod_logger

 # 获得当前路径下的文件名
def getFileName(suffix):
    fileNameList = [fileName for fileName in os.listdir('.') if fileName.rpartition('.')[2] == 'xls' or fileName.rpartition('.')[2] == 'xlsx'];

    fileNumList = [];
    if len(fileNameList) == 0:
        print("当前目录下没有xls格式的文件");
        return;
    elif len(fileNameList) > 1:
        numSuffix = -1; 
        fileSelect = "";
        fileTag = "";
        # 依次取出每个Excel文档。
        for file in fileNameList:
            # 去掉后缀
            fileNoSuffix = file.rpartition('.')[0];
            # 依照_分割成数组
            fileSplit = fileNoSuffix.split('_');
            # 如果只有两部分同时，第二部分为suffix标志
            if len(fileSplit) == 2 and fileSplit[1] == suffix:
                fileSelect = file;
                numSuffix = -1 ;
                fileTag = fileSplit[1];
            # 如果有三部分且含有数字，将数字记在fileNumList里面，用于后面比大小。
            elif len(fileSplit)  == 3:
                if int(fileSplit[2]) > numSuffix:
                    numSuffix = int (fileSplit[2]);
                    fileSelect = file;
                    fileTag = fileSplit[1];
            else :
                continue
    else :
        fileSelect = fileNameList[0];
        numSuffix = 0;
        fileNoSuffix = fileSelect.rpartition('.')[0];
        # 依照_分割成数组
        fileSplit = fileNoSuffix.split('_');
        fileTag = fileSplit[1];

    return fileSelect , numSuffix ,fileTag ;

    
MODEL= mod_config.getConfig('mode', 'model');
dbName = mod_config.getConfig('mode','db');
    
suffix = '含标准价'
[excelName , numSuffix , fileTag] = getFileName(suffix);

if fileTag.isdigit() and  MODEL == 'H3C' :
    inputKeys = ['ID','BOM','typeID','description','num','listprice','off','price','totalPrice','totalListPrice','productLine','waston','addOn'];
    outputKeys = ['ID','BOM','typeID','description','num','totalNum','listprice','off','price','totalPrice','totalListPrice','productLine','waston', 'addOn'];
    outputHeaderLine = ['序号','产品编码','产品型号','项目名称','数量',	'总数量','目录价','折扣','单价'	,'总价','总目录价','产线','WATSON_LINE_ITEM_ID','备注']
    suffix = '含标准价'
elif fileTag == suffix and MODEL == 'H3C':
    inputKeys = ['ID','BOM','typeID','description','num','totalNum','listprice','off','price','totalPrice','totalListPrice','productLine','waston', 'addOn'];
    outputKeys = ['ID','BOM','typeID','description','num','totalNum','listprice','off','price','totalPrice','totalListPrice','productLine','waston', 'addOn'];
    outputHeaderLine = ['序号','产品编码','产品型号','项目名称','数量',	'总数量','目录价','折扣','单价'	,'总价','总目录价','产线','WATSON_LINE_ITEM_ID','备注']
elif fileTag.isdigit() and MODEL == 'HPE':
    inputKeys = ['ID','BOM','typeID','description','num','listprice','off','price','totalPrice','totalListPrice','productLine','waston', 'addOn'];
    outputKeys = ['ID','num','BOM','typeID','description','totalNum','listprice','off','price','totalPrice','totalListPrice','productLine','waston', 'addOn'];
    outputHeaderLine = ['序号','单套数量','产品编码','产品型号','项目名称',	'总数量','目录价','折扣','单价'	,'总价','总目录价','产线','WATSON_LINE_ITEM_ID','备注']  
    
else:
    inputKeys = ['ID','num','BOM','typeID','description','totalNum','listprice','off','price','totalPrice','totalListPrice','productLine','waston', 'addOn'];
    outputKeys = ['ID','num','BOM','typeID','description','totalNum','listprice','off','price','totalPrice','totalListPrice','productLine','waston', 'addOn'];
    outputHeaderLine = ['序号','单套数量','产品编码','产品型号','项目名称',	'总数量','目录价','折扣','单价'	,'总价','总目录价','产线','WATSON_LINE_ITEM_ID','备注']  
    
if dbName == 'merchants':
    outputKeys = ['ID','H3C','BOM','Category','switch','typeID','description','totalNum','unit','price','totalPrice','taxCategory','taxRate','addOn','period','repairFree','maintenancePeriod','serialNum', 'productLine','waston'];
    # outputKeys = ['ID','H3C','BOM','Category','switch','typeID','description','totalNum','unit','listprice','totalListPrice','taxCategory','taxRate','addOn','period','repairFree','maintenancePeriod','serialNum', 'productLine','waston'];
    outputHeaderLine = ['序号','品牌','招银采购管理系统产品编码','产品分类','产品及服务名称','型号','详细描述（规格/技术参数）','数量','单位','单价','小计','纳税类别','税率','备注','到货周期（单位：工作日）','新购产品免费维护期（单位：月）','维护周期','维护产品序列号', '产线','WATSON_LINE_ITEM_ID']  
    

newSuffix = suffix + '_'+ str(numSuffix + 1);

excelTools = excelTools(3,excelName , inputKeys,outputKeys,outputHeaderLine,'价格明细清单',newSuffix)

excelTools.transToStandard();
# 增加标签，用来区分标题和普通报价
excelTools.addTagColumn();
# excelTools.removeOtherLines();
excelTools.getSubTotalIndex();

# 替换光模块
# excelTools.replaceSFP();

# 从数据库里面获取信息
listDB = excelTools.getValueByDB('typeID');
excelTools.replaceByList (listDB,excelTools.dbKeys);
# 增加公式
excelTools.addFormula();
# 所有的OFF都和总计行的OFF相等
if 'off' in outputKeys:
    excelTools.replaceOff();

# 标题用outputValues替换
excelTools.replaceTopRow();
# 删除BOM编码
# excelTools.removeBOM();
# 删除ID
excelTools.removeID();
# 打印明细sheet
if MODEL == 'HPE':
    hideColumn = ['productLine','waston','typeID']
else:
    hideColumn = ['productLine','waston'];

if dbName == 'merchants' :
   excelTools.replaceTotalNum();
   
excelTools.printList(excelTools.excelList , excelTools.sheetName , hideColumn);
# 打印sumary
excelTools.sumaryList = excelTools.getSumaryList();

if dbName != 'merchants':
    excelTools.replaceTotalNum();
# 打印汇总页
if MODEL == 'H3C':
    [bomDict , indexList]= excelTools.getCheckList('typeID');
else:
    [bomDict , indexList]= excelTools.getCheckList();

if dbName != 'merchants':
    excelTools.printCheckList(bomDict , indexList);
    
excelTools.printSumaryList();
excelTools.xlsWriterTools.closeWriter();