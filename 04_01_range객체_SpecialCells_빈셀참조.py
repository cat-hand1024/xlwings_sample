import xlwings as xw
from xlwings.constants import CellType
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='003_빈셀처리.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)

sht=wb.sheets(1)

cellTypeContains=CellType()

# VBA SpecialCells 빈셀 선택

r=sht.range('b5').current_region # 연속영역 선택
emptyCells=r.api.SpecialCells(cellTypeContains.xlCellTypeBlanks) # win32com에서 제공하는 VBA 모듈 적용 / !! 선택된 range는 win32com range임 !!
xlwingsRange=sht.range(emptyCells.GetAddress()) # 변환 ( win32com -> xlwings )
xlwingsRange.select()
xlwingsRange.value=0 # 공백을 0으로 변환
