import xlwings as xw
from xlwings.constants import CellType,AutoFilterOperator
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='010_복사ToFile.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

choice_R=sht.range('b5').current_region
print(choice_R.address)
# 필터링
# Range.api.AutoFilter( Field= , Criteria1=[(Criteria2=,,,,Operator= )])
choice_R.api.AutoFilter(Field=2,Criteria1='주임',Criteria2='사원',Operator=AutoFilterOperator.xlOr)

# 필터링된 화면만 복사
choice_R.api.SpecialCells(CellType.xlCellTypeVisible).Copy() # win32com 메서드 사용
sht.range('q14').paste(paste='formats')
sht.range('q14').paste(paste='column_widths')
sht.range('q14').paste(paste='values')