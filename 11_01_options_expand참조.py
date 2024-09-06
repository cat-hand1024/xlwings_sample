import xlwings as xw
import os


filePath=r'C:\코딩학습\xlwings\예제'
fileName='001_참조.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

r=sht.range('c6').options(expand='right').value
print(r)
r=sht.range('c6').options(expand='down').value
print(r)
r=sht.range('c6').options(expand='table').value
print(r)