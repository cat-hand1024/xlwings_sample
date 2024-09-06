import xlwings as xw
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='001_참조.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)


origin_R=sht.range('b5').current_region

origin_R.copy()

sht.range('n5').paste(transpose=True)