import xlwings as xw
import os


filePath=r'C:\코딩학습\xlwings\예제'
fileName='011_clear.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

r=sht.range('a1').current_region
r=r.offset(2,0).resize(r.rows.count-2,None)

allClear_R=r.columns(1)
allClear_R.clear()

contentsClear_R=r.columns(2)
contentsClear_R.clear_contents()

formatsClear_R=r.columns(3)
formatsClear_R.clear_formats()