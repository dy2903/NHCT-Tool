# !/usr/bin/env python3 
# -*-coding:UTF-8-*-
import xlsxwriter
import os.path
import mod_config
import mod_logger
import re
import mod_config
import mod_logger

MODEL= mod_config.getConfig('mode', 'model');
class xlsWriterTools:
    def __init__(self,list,outKeys,destFile,sheetName=u'价格明细清单'):
        self.excelList = list;
        self.headerTag = outKeys;
        self.destFile = destFile;
        self.sheetName = sheetName
        # 获得要输出的excel的文件流
        self.workbook = xlsxwriter.Workbook(self.destFile)
        # 获得格式
        [self.money , self.headerColor,self.subtotalColor,self.subtotalOff,self.totalColor,self.common , self.link_format , self.headerOff] = self.getFormat();
        # 获得excel表格的列和headerTag的关联
        self.columnIndexDict = self.getColumnIndexDict(self.headerTag);
        
        
    # 获得excel表格的列和headerTag的关联
    def getColumnIndexDict(self , indexList):
        columnIndexDict = {};
        columnTags = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];
        for i in range(len(indexList)):
                columnIndexDict[indexList[i]] = columnTags[i];

        return columnIndexDict;        
        
    # 获得交集
    def getIntersection(self , A , B):
        intersectionList = [i for i in A if i  in B];
        return intersectionList;
    
    # 设置sheetName页的列宽
    def setColumn(self,outKeys , sheetName=u'价格明细清单' , hideColumn = ['productLine','waston','typeID']):
        worksheet = self.workbook.get_worksheet_by_name(sheetName);
        breadthDict = {'ID':3,'BOM':10,'typeID':18,'description':45,'totalNum':5,'listprice':9,'off':10,'num':5,'price':9,'totalPrice':11,'totalListPrice':11,'addOn':13,'Description':15,'Quotation':30};
        # 输出列的对应关系
        columnIndexDict = self.getColumnIndexDict(outKeys);
        for tag in outKeys:
            if tag not in columnIndexDict:
                return;
                
            columnIndex = columnIndexDict[tag];
            # 默认设为10
            if tag not in breadthDict:
                breadthDict[tag]  = 10;
                
            breadth = breadthDict[tag];
            columnRange = columnIndex+':'+columnIndex;
            worksheet.set_column(columnRange , breadth  )
        
        # productLine 和waston列需要隐藏
        if self.getIntersection(outKeys , hideColumn) != []:          
            for tag in hideColumn:
                index = columnIndexDict[tag];
                breadth = breadthDict[tag];
                columnRange = index+':'+index;                
                worksheet.set_column(columnRange ,None, None, {'hidden': 1});
            
    # 设置格式
    def getFormat(self):
        if MODEL == 'HPE':
        # 普通行的格式
            common = self.workbook.add_format({'num_format':'0.00%' , 'font_name':'Arial','font_size':10} );
            # 显示列表价的格式
            money = self.workbook.add_format({'num_format': '#,##0' , 'font_name':'Arial','font_size':10})
            link_format = self.workbook.add_format({'color': 'blue', 'underline': 1 , 'font_name': 'Arial' , 'font_size':'10' })
            
        else:
            common = self.workbook.add_format({'num_format':'0.00%' , 'font_name':'Arial','font_size':10 , 'border':4} );
            # 显示列表价的格式
            money = self.workbook.add_format({'num_format': '#,##0' , 'font_name':'Arial','font_size':10,'border':4})            
            link_format = self.workbook.add_format({'color': 'blue', 'underline': 1, 'border':4 , 'font_name': 'Arial' , 'font_size':'10' })
        
        headerColor = self.workbook.add_format({'bg_color': '#339966','num_format': '#,##0','bold': True,'font_name':'Arial','font_size':10,'font_color':'#ffffff' })
        # 小计行格式
        subtotalColor = self.workbook.add_format({'bg_color': '#c0c0c0','num_format': '#,##0','bold': True,'bottom':1,'font_name':'Arial','font_size':10})
        # 总计行格式
        totalColor = self.workbook.add_format({'bg_color': '#c0c0c0','num_format': '#,##0','bold': True,'font_name':'Arial','font_size':10,'bottom':1})
        # 总计行的OFF的格式
        subtotalOff= self.workbook.add_format({'bg_color': '#c0c0c0','num_format': '#,##0','bold': True,'num_format':'0.00%' , 'font_name':'Arial','font_size':10,'bottom':1} );
        headerOff= self.workbook.add_format({'bg_color': '#339966','num_format': '#,##0','bold': True,'num_format':'0.00%' , 'font_name':'Arial','font_size':10,'bottom':1} );
        
        return  money , headerColor,subtotalColor,subtotalOff , totalColor,common , link_format ,headerOff;
        
    # 加新的一页    
    def addNewSheet (self , sheetName):
        worksheet = self.workbook.add_worksheet(sheetName);
        return worksheet;
        
    # 打印
    def printList (self,list,header,sheetName=u'价格明细清单'):
        if self.workbook.get_worksheet_by_name(sheetName) is None:
            worksheet = self.addNewSheet(sheetName);
        else :
            worksheet = self.workbook.get_worksheet_by_name(sheetName);
            
        col = 0;        
        listLength = len(list)        
        for i in range(listLength):        
            for outkey in header:
                if list[i]['colorTag'] == 'title':
                    if outkey == 'off':
                        worksheet.write(i, col, list[i][outkey], self.headerOff);
                        col += 1;
                    else:  
                        worksheet.write(i, col, list[i][outkey], self.headerColor);
                        col += 1;
                elif list[i]['colorTag'] == 'subtotal':
                    # 对于小计行，off单元格的格式单独设定
                    if outkey == 'off':
                        worksheet.write(i, col, list[i][outkey], self.subtotalOff);
                        col += 1;
                    else:    
                        worksheet.write(i, col, list[i][outkey], self.subtotalColor);
                        col += 1;
                        
                elif list[i]['colorTag'] == 'totalSum':
                    if outkey == 'off':
                        worksheet.write(i, col, list[i][outkey], self.subtotalOff);
                        col += 1;
                    else:                 
                        worksheet.write(i, col, list[i][outkey], self.totalColor);
                        col += 1;
                else:
                    if outkey == 'off'or outkey == 'rate' or outkey == 'addOn':
                        worksheet.write(i, col, list[i][outkey],self.common);
                        col += 1;
                    else:
                        worksheet.write(i, col, list[i][outkey],self.money);
                        col += 1;
            # 打印完一行回到开头            
            col = 0;

        return  worksheet;
    # 设置URL
    def writeURL(self , list , keys , sheetName):
        worksheet = self.workbook.get_worksheet_by_name(sheetName)
        columnIndexDict = {};
        columnIndexDict = self.getColumnIndexDict(keys);
        for i in range(1 , len(list) - 1):  
            if 'description' not in list[i].keys():
                print('没有描述列');
                

            cellIndex = columnIndexDict['description'] + str(i+1);
            worksheet.write_url(cellIndex , list[i]['description'] ,self.link_format, list[i]['descString']);
            
        return
            
        
    # 设置行宽
    def setRow (self , sheetName , list):
        worksheet = self.workbook.get_worksheet_by_name(sheetName);
        for i in range(len(list)):
            worksheet.set_row(i , 23 );
    
    # 设置筛选列
    def setAutofilter(self , sheetName , list):
        worksheet = self.workbook.get_worksheet_by_name(sheetName);
        begin = "A1";
        endColumn = self.columnIndexDict[self.headerTag[len(self.headerTag)-1]];
        endRow = len(list) ;
        end = endColumn + str(endRow);

        filterRange = begin +':'+ end;
        worksheet.autofilter(filterRange);
    # 冻结首行
    def freezeTopRow(self , sheetName):
        worksheet = self.workbook.get_worksheet_by_name(sheetName);
        worksheet.freeze_panes(1,0);
        
    # 关闭
    def closeWriter(self):
        self.workbook.close();
    