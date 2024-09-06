import xlwings as xw
from openpyxl.utils import get_column_letter, column_index_from_string
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='002_end.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

#  공백이 있는 영역 참조
startCell=sht.range('b5')
rightEndCol=startCell.end('right').column
sheetLastCell=sht.range(sht.cells.last_cell.row,rightEndCol)
rightBottomCell=sheetLastCell.end('up')

r=sht.range(startCell,rightBottomCell)
r.select()

