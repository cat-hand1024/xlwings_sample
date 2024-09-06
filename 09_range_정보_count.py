import xlwings as xw
import os


filePath=r'C:\코딩학습\xlwings\예제'
fileName='000_자료입력.xlsx'
file=os.path.join(filePath,fileName)


wb=xw.Book(file)
sht=wb.sheets(1)

r=sht.range('b1').current_region
print('행갯수 : ',r.rows.count)
print('컬럼갯수 : ',r.columns.count)