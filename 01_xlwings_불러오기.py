import xlwings as xw
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='001_참조.xlsx'

file=os.path.join(filePath,fileName)

wBook=xw.Book(file)

shtName=wBook.sheet_names
print(shtName)

shtCount=wBook.sheets.count
print(shtCount)

sht=wBook.sheets(1)
print(sht.name)
