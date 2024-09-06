import xlwings as xw
from xlwings.constants import CellType
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='003_빈셀처리.xlsx'
file=os.path.join(filePath,fileName)
xb=xw.Book(file)
sht=xb.sheets(1)

r=sht.range('b5').current_region

screen_WR=r.api.SpecialCells(CellType.xlCellTypeVisible)
screen_XR=sht.range(screen_WR.GetAddress())
print(screen_XR.address)
