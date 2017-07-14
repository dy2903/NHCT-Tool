# !/usr/bin/env python3 
# -*-coding:UTF-8-*-
# from openpyxl import Workbook
# from openpyxl import load_workbook
import xlrd

# wb = load_workbook("./【深证通】H3C配置清单-2017-7-3_含标准价.xlsx")
wb = xlrd.open_workbook("./中国太平保险集团公司武汉数据中心网络设备采购项目_含标准价.xls" )
ws = wb.sheet_by_name('价格明细清单')
rowValues = ws.row_values(4)
print(rowValues)
# tuple(ws['A1':'C3'])
# #
# for rowOfCellObjects in ws['A1':'F3']:
#     for cellObj in rowOfCellObjects:
#         print(cellObj.coordinate, cellObj.value);
#
# print(ws.cell(row=4, column=5).value)
# for rowOfCellObjects in ws['A1':'F3']:
#     for cellObj in rowOfCellObjects:
#         print(cellObj.coordinate, cellObj.value);
#
# print(ws.cell(row=4, column=5).value)