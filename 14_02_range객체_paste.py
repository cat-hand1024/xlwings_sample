import xlwings as xw
import os


filePath=r'C:\코딩학습\xlwings\예제'
fileName='007_셀복사.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

# 3. clipborder 복사 (b3:d3 -> clipborder ->paste ->복사모드해제)
sht.range('b3,d3').copy()
sht.range('m3').paste()
sht.range('m4').paste()
wb.api.Application.CutCopyMode=False # 엑셀복사모드 해제

