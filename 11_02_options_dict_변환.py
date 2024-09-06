import xlwings as xw
import os


filePath=r'C:\코딩학습\xlwings\예제'
fileName='013_OptionsToDict.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

r=sht.range('b5').current_region
r=r.offset(1,0).resize(r.rows.count-2,None)
dictRange=sht.range(r.columns(1),r.columns(2))
v=dictRange.options(dict).value
print(v)
