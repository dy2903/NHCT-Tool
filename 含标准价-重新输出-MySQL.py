#/usr/bin/env python3
# -*- coding:UTF-8-*-
from quotationTools.excelTools  import excelTools
import os.path
import re

 # 获得当前路径下的文件名
def getFileName(suffix):
    # fileNameList = [];
    fileNameList = [fileName for fileName in os.listdir('.') if fileName[-3:] == 'xls'];
    
    if len(fileNameList) == 0:
        print("当前目录下没有xls格式的文件");
        return;
    elif len(fileNameList) > 1:
        print("当前目录下有多余一个的xls格式的文件");
        for file in fileNameList:
            # if re.match('\w*\_\d*',file):
            if file.partition('_')[2][:-4] == suffix:
                return file[:-4];
                
        return;
    else:
        return fileNameList[0][:-4];


   
inputKeys = ['ID','BOM','typeID','description','num','totalNum','off','listprice','price','totalPrice','totalListPrice','productLine','waston', 'addOn'];
outputKeys = ['ID','BOM','typeID','description','num','totalNum','off','listprice','price','totalPrice','totalListPrice','productLine','waston', 'addOn'];

# outputKeys = ['ID','BOM','typeID','description','num','totalNum','off','listprice','price','totalPrice','totalListPrice','productLine','waston', 'addOn'];

outputHeaderLine = ['序号','产品编码','产品型号','项目名称','数量',	'总数量','折扣','目录价','单价'	,'总价','总目录价','产线','WATSON_LINE_ITEM_ID','备注']

suffix = '含标准价'
excelName = getFileName(suffix);
if excelName.partition('_')[2] == suffix:
    suffix = '含标准价_2'
    
excelTools = excelTools(3,excelName , inputKeys,outputKeys,outputHeaderLine,'价格明细清单',suffix)

excelTools.transToStandard();
# 增加标签，用来区分标题和普通报价
excelTools.addTagColumn();
# excelTools.removeOtherLines();
excelTools.getSubTotalIndex();
# 替换光模块
# excelTools.replaceSFP();

# 从数据库里面获取信息
listDB = excelTools.getValueByDB();
excelTools.replaceByList (listDB,excelTools.dbKeys);
# 增加公式
excelTools.addFormula();
# 所有的OFF都和总计行的OFF相等
excelTools.replaceOff();
# 标题用outputValues替换
excelTools.replaceTopRow();
# 删除BOM编码
# excelTools.removeBOM();
# 删除ID
excelTools.removeID();
# 打印明细sheet
excelTools.printList(excelTools.excelList , excelTools.sheetName);
# 打印sumary
excelTools.sumaryList = excelTools.getSumaryList();
# 打印汇总页
[bomDict , indexList]= excelTools.getCheckList();
excelTools.printCheckList(bomDict , indexList);
excelTools.printSumaryList();
excelTools.xlsWriterTools.closeWriter();