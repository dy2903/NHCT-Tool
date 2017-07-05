# !/usr/bin/env python3 
# -*-coding:UTF-8-*-
import os.path

import xlrd

import mod_config
import mod_logger
# 引入类
from mysqlTools.MySQLDBTools import database
from quotationTools.xlsWriterTools import xlsWriterTools

# 导入配置
# 日志路径、文件路径
LOGPATH = mod_config.getConfig('path', 'logpath') + 'database.log'
EXCELPATH = mod_config.getConfig('path', 'excelPath');
EXCELFILENAME = mod_config.getConfig('path', 'excelName');
MODEL= mod_config.getConfig('mode', 'model');
# 初始化日志类
logger = mod_logger.logger(LOGPATH)
class excelTools:
    # inputKeys：输入的标志符，用于数组的标识
    # outKeys：要输出的列的标识
    # outputValues：要输出的表的标题行
    # suffix:后缀：含标准价，不含标准价
    def __init__(self, titleRowIndex,excelName ,  inputKeys,outKeys,outputValues,sheetName,suffix):
        self._logger = logger;
        excelPath = EXCELPATH;
        # self._excelName = self.getFileName()[:-4];
        self._excelName = excelName;
        # self._excelPathName = os.path.join(excelPath, self._excelName + '.xls');
        self._excelPathName = os.path.join(excelPath, self._excelName );
        self.sheetName = sheetName;
        self.inputKeys = inputKeys;
        self.outKeys = outKeys;
        self.outputValues = outputValues;
        # 获得文件流
        self._workbookStream = self.openExcel();
        splitName = self._excelName.partition('_')[0];
        # 目的xls的地址
        self._destFile = os.path.join(excelPath, splitName + '_' + suffix+'.xls');
        # 比较要输出的列标题和输入的标题，得出其差集
        self.diffList = self.getDiffList();
        # 把excel的一个sheet转换为数组，形式为:[{'列表价' : 10000} , {'单价':100}]
        self.detailSheet = self._workbookStream.sheet_by_name(self.sheetName);
        # 判断是哪一种类型的表格
        if self.detailSheet.row_values(0)[1] != '':
            titleRowIndex = 0;
        elif self.detailSheet.row_values(4)[0] == '序号':
            titleRowIndex = 4;
        else :
            pass;
        [self.excelList, self.headerLine] = self.getDictList(titleRowIndex);
        self.listLength = len(self.excelList);
        # 输出的列对于Excel表格里面的列
        self.columnIndexDict = self.getColumnIndexDict(self.outKeys);
        # outKeys与输出的表头的对应关系
        self.outputKeyDict = self.getOutputKeyDict();
        # 输入Excel的工具类
        self.xlsWriterTools = xlsWriterTools(self.excelList,self.outKeys,self._destFile,self.sheetName);
        # Summary sheet的key
        self.sumaryListKeys = ['ID','description','Quotation','Qty','price','totalPrice','rate'];
        # Sumary sheet 的标题
        self.sumaryValues = ['序号','描述','配置主机','数量','单价','总价','占比'];
        # 数据库的列标题
        self.dbKeys = ['ID', 'BOM', 'typeID', 'description', 'listprice'];
        
        
    # 获得差集
    def getDiffList(self):
        diffList = [i for i in self.outKeys if i not in self.inputKeys];
        return diffList;
        
        
        
    # 打开表格
    def openExcel(self):
        try:
            workbookStream = xlrd.open_workbook(self._excelPathName)
            # wb =  workbookStream.sheet_by_name('价格明细清单');
            # values = wb.row_values(3);
            # print(values)
            return workbookStream
        except Exception as data:
            logger.error("Can't Open the excel, %s" % data)
            
            
    # /*将Excel读出，并转换为数组的形式，数组的每一行都是dict形式*/
    def getDictList(self, initialLineIndex=0 ):
        nrows = self.detailSheet.nrows  # 行数
        # colnameindex指的是标题列所在的位置，注意是从0开始的
        headerLine = self.inputKeys;  # 标题栏对应的key
        # excel表格中每一行数据都对应list中的一行，每一行数据都是一个dict，键为当前列的标题
        list = []
        # range从0开始，需要包含标题行
        for i in range(initialLineIndex, nrows):
            # 取出每一行的数据
            rowValues = self.detailSheet.row_values(i)
            # print(rowValues)
            # 如果不为空的话，组装成dict
            #  形式为:[{'列表价' : 10000} , {'单价':100}]
            if rowValues:
                rowDict = {}
                for j in range(len(headerLine)):
                    # 将某一行某一列的数据rowValues[j]首先作为键值与当前的key colnames[j]（也就是当前列标题）组合成为dict，然后附加到list的末尾。
                    rowDict[headerLine[j]] = rowValues[j]
                list.append(rowDict)
        return list, headerLine

        
        
    # 新加一列
    def addNewColumn (self, key , value):
        self.inputKeys.append(key);
        for i in range(len(self.excelList)):
            self.excelList[i][key] = value;
      

      
    # 删除标题下的空行
    def deleteEmptyRow(self):
        if self.excelList[1]['price'] == '' and self.excelList[1]['ID'] == '':
            del self.excelList[1];
     
     
    # 获得新的一行，并且在里面加入相应的值        
    def getEmptyLine(self , key , value):
        dict = {};
        for i in range(len(self.excelList[0])):
            dict[list(self.excelList[0].keys())[i]] = '';
            
        dict[key] = value;
        return dict;        
            
    # /**对读出的列表进行调准，增加可能没有的小计行**/
    def transToStandard(self):
        # 删除空行
        self.deleteEmptyRow();
        # 增加总数量行
        for i in self.diffList:
            self.addNewColumn(i,'');

        self.addNewColumn('colorTag' , 'common')
        # 如果没有小计行，则增加
        for i in range(len(self.excelList) - 1, 1, -1):
            if self.excelList[i]['ID'] != '' and self.excelList[i - 1]['BOM'] != '小计':
                dict1 = self.getEmptyLine('BOM','小计');
                self.excelList.insert(i, dict1);
            elif self.excelList[i]['BOM'].find('#') != -1 and self.excelList[i]['ID'] == ''and self.excelList[i]['typeID'] == '':
                self.excelList.pop(i);
            elif self.excelList[i]['description'] == 'Factory integrated':
                self.excelList.pop(i);
            else:
                continue;

        if self.excelList[1]['ID'] == "":
            tmp = self.getEmptyLine('ID',1);
            self.excelList.insert(1, tmp);
            
        if self.excelList[len(self.excelList) - 1]['BOM'] != '总计':
            dict = self.getEmptyLine('BOM','总计');
            self.excelList.append(dict);
            
        if self.excelList[len(self.excelList) - 2]['BOM'] != '小计':
            dict2 = self.getEmptyLine('BOM','小计');
            self.excelList.insert(len(self.excelList) - 1, dict2);
            
     
     
    # 打上标记用来区分标题行、小计行以及普通的报价行   
    def addTagColumn(self):
        # 从后往前进行遍历
        for i in range(len(self.excelList) - 1, -1, -1):
            if self.excelList[i]['ID'] != '' and i != 0:
                self.excelList[i]['colorTag'] = 'title';
            elif self.excelList[i]['BOM'] == '小计' :
                self.excelList[i]['colorTag'] = 'subtotal';
            elif self.excelList[i]['BOM'] == '总计':
                self.excelList[i]['colorTag'] = 'totalSum';
            else:
                    self.excelList[i]['colorTag'] = 'common';
        return;
        
    # /**获取小计行，标题行的行号**/    
    def getSubTotalIndex (self):
        headerIndex = [];
        subtotalIndex = [];
        totalSumIndex = 0;
        for i in range(len(self.excelList)):
            rowDict = self.excelList[i];
            if rowDict['colorTag'] == 'title':
                headerIndex.append(i);
            elif rowDict['colorTag'] == 'subtotal':
                subtotalIndex.append(i);
            elif rowDict['colorTag'] == 'totalSum':
                totalSumIndex = i;
            else:
                continue;

        self.excelList[0]['colorTag'] = 'subtotal';
        
        self.headerIndex = headerIndex;
        self.subtotalIndex = subtotalIndex;
        self.totalSumIndex = totalSumIndex;
    
    
    def mainframeFormula(self,i ,tag):
        if (tag in self.inputKeys) == 0 or (i < 0):
            print('tag is not in inputKeys');
            return;
            
        self.excelList[self.subtotalIndex[i]][tag] = '=' + self.columnIndexDict[tag] + str(self.headerIndex[i] + 2);
    
    # 小计行的公式
    def subtotalFormula(self, i, tag):
        if (tag in self.inputKeys) == 0 or (i < 0):
            print('tag is not in inputKeys');
            return;
            
        subtotalListIndex = self.subtotalIndex;
        begin = self.headerIndex[i] + 2;
        end = subtotalListIndex[i]
        rowIndex = self.columnIndexDict[tag];
        
        self.excelList[end][tag] = '=SUM(' + rowIndex + str(
            begin) + ':' + rowIndex + str(end) + ')';
    
    # 价格公式
    def priceFormula(self, i, tagOut, tagIn1, tagIn2):
        if (tagOut in self.inputKeys) == 0 or tagIn1 in self.inputKeys == 0 or tagIn2 in self.inputKeys == 0 or (i < 0):
            print('tag is not in inputKeys');
            return;
        self.excelList[i][tagOut] = '=' + self.columnIndexDict[tagIn1] + str(i + 1) + '*' + self.columnIndexDict[tagIn2] + str(i + 1);
    
    # 最终价格
    def amountFormular(self, i, tag):
        if (tag in self.inputKeys) == 0 or (i < 0):
            print('tag is not in inputKeys');
            return;
        # le = self.totalSumIndex ;
        amountIndex = len(self.excelList) - 1;
        self.excelList[i][tag] = '=SUM(' + self.columnIndexDict[tag] + '2:' +self.columnIndexDict[tag] + str(amountIndex) + ')/2';
    
    # 每一套的数量 * 套数
    def addTotalNumFormular(self ,sectionNum):           
        if 'num' in self.outKeys:
            for i in range(len(self.headerIndex)):
                    if self.excelList[self.headerIndex[i]]['num'] =='' :
                        self.excelList[self.headerIndex[i]]['num'] = sectionNum[i];

                    begin = self.headerIndex[i] + 1; 
                    end = self.subtotalIndex[i];
                    for j in range(begin,end):
                        self.totalNumFormular(i,j,'num');
        else:
            for j in range(len(self.excelList)):
                self.excelList[j]['totalNum'] = self.excelList[j]['num'];
        
    def totalNumFormular(self , i,j , tag):
        self.excelList[j]['totalNum'] = '=$' + self.columnIndexDict[tag] + '$' + str(self.headerIndex[i] + 1) + '*' + self.columnIndexDict[tag] + str(j + 1);
        
    # 小计行的公式
    def addSubtotalFormula(self,tag):
        subtotalIndex = self.subtotalIndex;
        if len(subtotalIndex) <= 0 :
            return;
        
        for i in range(len(subtotalIndex)):
            self.subtotalFormula(i, tag);
            self.mainframeFormula(i,'typeID');
        
    # 每一行价格单元格的公式    
    def addPriceFormula(self , off = 0.25):
        for i in range(len(self.excelList)):
            if self.excelList[i]['colorTag'] == 'common':
                if 'listprice' in self.outKeys:
                    self.priceFormula(i, 'price', 'off', 'listprice');
                    self.priceFormula(i, 'totalListPrice', 'listprice', 'totalNum');
                    self.priceFormula(i, 'totalPrice', 'price', 'totalNum');
                else:
                    # self.excelList[i]['price'] = off * self.excelList[i]['listprice'];
                    self.priceFormula(i, 'totalPrice', 'price', 'totalNum');   
    
    # 总计
    def addAmountFormula(self)    :
        self.amountFormular(len(self.excelList)-1, 'totalPrice');
        if 'totalListPrice' in self.outKeys:
            self.amountFormular(len(self.excelList) - 1, 'totalListPrice');
            # index = len(self.excelList) - 1;
            # tag = 'totalListPrice'
            # amountIndex = len(self.excelList) - 1;
            # self.excelList[index][tag] = '=SUMPRODUCT((' + self.columnIndexDict[tag] + '2:' +self.columnIndexDict[tag] + str(amountIndex) + ')/2';
            # =SUMPRODUCT((E2:E24<>"")*(K2:K24))
    # 加所有的公式
    def addFormula (self):
        sectionNum = [1] * len(self.headerIndex);    
        self.addTotalNumFormular(sectionNum);
        self.addSubtotalFormula('totalPrice');
        if 'totalListPrice' in self.outKeys:
            self.addSubtotalFormula('totalListPrice');

        self.addPriceFormula();
        self.addAmountFormula();

    # 获得每一列所有的的列号
    def getColumnIndexDict(self , columnIndex):        
        columnIndexDict = {};
        columnTags = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];
        try:
            for i in range(len(columnIndex)):
                columnIndexDict[columnIndex[i]] = columnTags[i];
        except Exception as data:
            self._logger.error('titleTags比columnTags长:%s' % data);

        return columnIndexDict;
    
    # 将outputKeys和outputValues关联起来
    def getOutputKeyDict(self ):
        if len(self.outKeys) != len(self.outputValues):
            print("输入的list长度不一致");
            return;
            
        tagsMapHeader = {};
        for i in range(len(self.outKeys)):
            tagsMapHeader[self.outKeys[i]] = self.outputValues[i];
        return tagsMapHeader;
    
    
    # 原来的标题行用自定义的标题替换
    def replaceTopRow (self):
        for key in self.outKeys:
            self.excelList[0][key] = self.outputKeyDict[key];
            
        return
    
    # 打印清单sheet
    def printList (self , excelList , sheetName , hideColumn = ['productLine','waston','typeID']):
        # 打印清单sheet
        self.xlsWriterTools.printList(excelList , self.outKeys , sheetName);
        # 设置每一列的格式
        self.xlsWriterTools.setColumn(self.outKeys , sheetName ,hideColumn);
        # 设置每一行的格式
        if MODEL == 'H3C' :
            self.xlsWriterTools.setRow( sheetName , excelList);
        # 设置自动过滤
        self.xlsWriterTools.setAutofilter( sheetName ,excelList);
        # 冻结首行
        self.xlsWriterTools.freezeTopRow( sheetName);
        # 暂时不关闭xlsxWriter
        # self.xlsWriterTools.closeWriter();
    
    
    # 去掉BOM编码    
    def removeBOM(self):
        if 'BOM' not in self.inputKeys:
            # print("没有BOM列");
            self.addNewColumn('BOM',"");
            
        for i in range(1,len(self.excelList)):
            row = self.excelList[i];
            if row['colorTag'] == 'common':
                row['BOM'] = '';    
                
    # 去掉BOM编码    
    def removeID(self):
        if 'ID' not in self.inputKeys:
            # print("没有BOM列");
            self.addNewColumn('ID',"");

        titleCount = 1;
        for row in self.excelList:
            if row['colorTag'] == 'title':
                row['ID'] = titleCount;
                titleCount += 1;
            else:
                row['ID'] = '';

    # SFP-XG-SX-MM850-E全部转换为不带-E的模块        
    def replaceSFP(self):
        if 'typeID' not in self.inputKeys:
            # print("没有BOM列");
            self.addNewColumn('typeID',"");
            
        for i in range(1,len(self.excelList)):
            row = self.excelList[i];
            if row['colorTag'] == 'common' and row['typeID'] == "SFP-XG-SX-MM850-E":
                row['typeID'] = 'SFP-XG-SX-MM850';   
    
    # 获得Sumary 页的数组
    def getSumaryList (self):
        sumaryList = [];
        # 首行
        dictTMP = {};
        for  i in range(len(self.sumaryListKeys)):
            sumaryListKey = self.sumaryListKeys[i];
            dictTMP[sumaryListKey] = self.sumaryValues[i];
        
        dictTMP['colorTag'] = 'subtotal';
        dictTMP['descString'] = '描述';
        sumaryList.append(dictTMP);
        
        # 设置每一个单元格的公式
        # 每一个header列为一行
        for i in range(len(self.headerIndex)):
            dict = {};
            header = self.headerIndex[i]
            dict ['Qty'] = '='+self.sheetName+'!'+self.columnIndexDict['num'] + str(header + 1) ;
            # URL连接
            dict['description'] = 'internal:'+self.sheetName+'!'+self.columnIndexDict['BOM'] + str(header + 1) ;
            # 描述的值
            dict['descString'] = self.excelList[header]['BOM'];
            # self.subtotalIndex表示的是明细清单中的每个项目小计行的索引
            subtotal = self.subtotalIndex[i];
            dict['Quotation'] = '='+self.sheetName+'!'+self.columnIndexDict['typeID'] + str(subtotal + 1);
            dict['totalPrice'] = '='+self.sheetName+'!' + self.columnIndexDict['totalPrice'] + str(subtotal + 1);
            dict['price'] = '=F'+str(i+2)+'/D'+str(i+2);
            dict['colorTag'] = 'common';
            dict['ID'] = i + 1;
            # 每个项目的占比
            dict['rate'] = '=F' + str(i+2) + '/'  +self.sheetName+'!' + self.columnIndexDict['totalPrice'] + str( self.totalSumIndex+ 1);
            sumaryList.append(dict);

        dictTMP = {};
        for  i in range(len(self.sumaryListKeys)):
            sumaryListKey = self.sumaryListKeys[i];
            dictTMP[sumaryListKey] = '';
        
        dictTMP['colorTag'] = 'subtotal';
        dictTMP['descString'] = '';
        dictTMP['totalPrice'] = '=SUM(F2:' +  'F' + str (len(sumaryList)) + ')';
        sumaryList.append(dictTMP);
        
        
        return sumaryList;
        
        
     # 打印Sumary页   
    def printSumaryList (self ):
        # 新加一页，注意要使用原来的xlsWriterTools
        sumarySheet = self.xlsWriterTools.printList(self.sumaryList ,self.sumaryListKeys , 'Summary' );
        # 设置URL，链接到明细页的header行
        self.xlsWriterTools.writeURL(self.sumaryList , self.sumaryListKeys , 'Summary');
        # 设置列宽
        self.xlsWriterTools.setColumn(self.sumaryListKeys,'Summary');
        # 设置行宽
        self.xlsWriterTools.setRow('Summary' , self.sumaryList);
        # 关闭
        # self.xlsWriterTools.closeWriter();

    # 使用数据库来获得值
    def getValueByDB(self , tag='BOM'):
        self.db = database();
        
        listDB = [];
        dbKeys = self.dbKeys;
        for i in range( len(self.excelList)):
            dict = {};
            # 依次取出每一行
            row = self.excelList[i];
            if row['BOM'] == ''  and row['typeID'] != '':
                tag = 'typeID';
            elif row['BOM'] != '':
                tag = 'BOM';
            else :
                print('BOM 和产品型号都为空');
                
            sql = "SELECT *  from listpricetable where " + tag + "=%s";
            values = self.db.fetch_all(sql , row[tag]);
            # print (values);
            # 如果查出来的结果为空，则保留原样
            if values == []:
                for index in range(len(dbKeys)):
                    key = dbKeys[index];
                    values.append(row[key]);

                for j in range(len(dbKeys)):
                    dict[dbKeys[j]] = values[j];
            else:
                for j in range(len(dbKeys)):
                    dict[dbKeys[j]] = values[0][j];
                
            listDB.append(dict);
         # 关闭数据库   
        self.db.close();
        return listDB;
    
    # 替换self.excelList的某几列
    def replaceByList(self , listDB , keys ):
        if len(listDB) != len(self.excelList):
            print("替换的数组行数应该是原数组一样");
            
        for i in range(len(self.excelList)):
            row = self.excelList[i];
            for key in keys :
                row[key] = listDB[i][key];
            
        return;
    
    # 为了能获得归并以后的sheet，先把header行删除
    def removeOtherLines(self):
        removedList = self.excelList.copy();
        for i in range(len(removedList) - 3 , 1 , -1 ):
            row = removedList[i]
            if row['colorTag'] == 'title' or row['colorTag'] == 'subtotal':
                del removedList[i]

        del removedList[1]
                
        return removedList
    
    
    def replaceTotalNum (self):
        # i表示每个header行的index
        for i in range(len(self.headerIndex)):
            begin = self.headerIndex[i] + 1 ; 
            end = self.subtotalIndex[i];
            for j in range(begin,end):
                row = self.excelList[j];
                row['totalNum'] = self.excelList[self.headerIndex[i]]['num'] * self.excelList[j]['num'];
              
                        
    # 获得归并页
    def getCheckList(self , type='BOM'):
        # 如果不用这种复制的方法，所有的修改都会体现到原有的数组中
        removedList = self.removeOtherLines().copy();
        bomDict = {};
        indexList = [];
        for row in removedList:
            # 每一行的BOM值作为key，本行的所有信息作为值构成一个dict
            bom = row[type];
            # 如果第一次出现，则新建
            if bom not in bomDict.keys ():
                bomDict [bom] = row ;
                indexList.append(bom);
            else:
                # 其他的追加
                bomDict[bom]['totalNum'] += row['totalNum'];
                # bomDict[bom]['num'] += row['num'];
        
        return bomDict , indexList;
                
    def printCheckList (self , bomDict , bomList): 
        checkList = [];
        for bom in bomList:
            checkList.append(bomDict[bom]);
        # 删除剩下的小计行
        del checkList[-2];
        # 从表头下的一行开始
        # for i in range(2,len(checkList)-1):
        for i in range(1,len(checkList)-1):
            row = checkList[i];
            # if 'num' in self.inputKeys:
            # # 总数量和原有的数量保持一致
            #     row['totalNum'] = row['num'];
            
            # 与明细页进行联动
            row['off'] = '=价格明细清单!'+  self.columnIndexDict['off'] + str(self.totalSumIndex + 1)                
            row['totalListPrice'] = '=' + self.columnIndexDict['listprice'] + str(i + 1) + '*' + self.columnIndexDict['totalNum'] + str(i + 1);
            row['totalPrice'] = '=' + self.columnIndexDict['price'] + str(i + 1) + '*' + self.columnIndexDict['totalNum'] + str(i + 1);
            row['price'] = '=' + self.columnIndexDict['listprice'] + str(i + 1) + '*' + self.columnIndexDict['off'] + str(i + 1);
            row ['addOn'] = '=' + self.columnIndexDict['totalPrice'] + str(i + 1) + '/价格明细清单!'+ self.columnIndexDict['totalPrice'] + str( self.totalSumIndex+ 1) ;
        
        # 最后设置总价格的公式
        checkList[-1]['totalPrice'] = '=SUM(' + self.columnIndexDict['totalPrice']+'2:'+self.columnIndexDict['totalPrice']+ str(len(checkList) - 1) + ")";
        checkList[-1]['totalListPrice'] = '=SUM(' + self.columnIndexDict['totalListPrice']+'2:'+self.columnIndexDict['totalListPrice']+ str(len(checkList) - 1) + ")";
        checkList[-1]['off'] = '';
        
        
        self.printList(checkList , '归并页');
    
    # off和总计行的off保持一致
    def replaceOff (self) :
        for row in self.excelList:
            if row['colorTag'] == 'common':
                row['off'] = '=' + self.columnIndexDict['off'] + str(self.totalSumIndex + 1);
                
        for i in range(len(self.headerIndex)):
            self.excelList[self.headerIndex[i]]['off'] = '=' + self.columnIndexDict['off'] + str (self.totalSumIndex  + 1);
            begin = self.headerIndex[i] + 1; 
            end = self.subtotalIndex[i];
            for j in range(begin,end):
                self.excelList[j]['off'] = '=' + self.columnIndexDict['off'] + str(self.headerIndex[i] + 1);
        
        self.excelList[-1]['off'] = 1;
        
