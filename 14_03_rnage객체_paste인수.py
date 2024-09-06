import xlwings as xw
from openpyxl.utils import column_index_from_string
import os


filePath=r'C:\코딩학습\xlwings\예제'
fileName='014_Paste붙여넣기.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

sourceRange=sht.range('b2').current_region

sourceRange.copy() # clipborder에 복사

columnNum=column_index_from_string('f')

# 기본 복사
sht.range(1,columnNum).value='모두복사"all"'
sht.range(2,columnNum).paste(paste='all')

# 양식 붙여넣기 paste='formats'
columnNum=columnNum+3
sht.range(1,columnNum).value='형식복사"formats"'
sht.range(2,columnNum).paste(paste='formats')
#
# 값 붙여넣기 paste='value'
columnNum=columnNum+3
sht.range(1,columnNum).value='값복사"values"'
sht.range(2,columnNum).paste(paste='values')

# # 수식 붙여넣기 paste='formulas'
columnNum=columnNum+3
sht.range(1,columnNum).value='수식복사"formulas"'
sht.range(2,columnNum).paste(paste='formulas')

# # 값 & 숫자양식
columnNum=columnNum+3
sht.range(1,columnNum).value='수식복사"value_and_number_formats"'
sht.range(2,columnNum).paste(paste='values_and_number_formats')

#
# #테두리 제외
columnNum=columnNum+3
sht.range(1,columnNum).value='수식복사"all_except_borders"'
sht.range(2,columnNum).paste(paste='all_except_borders')

sht.api.Application.CutCopyMode=False
