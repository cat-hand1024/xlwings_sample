import xlwings as xw
import os


filePath=r'C:\코딩학습\xlwings\예제'
fileName='003_빈셀처리.xlsx'
file=os.path.join(filePath,fileName)


wb=xw.Book(file)
sht=wb.sheets(1)

originalRange=sht.range('자료영역')
originalRange.select()

# 총합계 제외
exceptTotalRowRange=originalRange.resize(originalRange.rows.count-1,None)
exceptTotalRowRange.select()
exceptTotalColRange=originalRange.resize(None,originalRange.columns.count-1)
exceptTotalColRange.select()
exceptTotalRange=originalRange.resize(originalRange.rows.count-1,originalRange.columns.count-1)
exceptTotalRange.select()
