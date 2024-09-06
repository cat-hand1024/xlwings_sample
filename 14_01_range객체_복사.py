import xlwings as xw
import os


filePath=r'C:\코딩학습\xlwings\예제'
fileName='007_셀복사.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)
#셀복사
target_Range=sht.range('f2')
sht.range('b2').copy(destination=target_Range) # b2의 내용을 f2로 복사
# 영역 복사
target_Range=sht.range('m2')
sht.range('b2').current_region.copy(destination=target_Range) # b2의 표를 전체 복사
