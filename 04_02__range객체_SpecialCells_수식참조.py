import xlwings as xw
import pandas as pd
from xlwings.constants import CellType
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='003_빈셀처리.xlsx'
file=os.path.join(filePath,fileName)
xb=xw.Book(file)
sht=xb.sheets(1)

r=sht.range('b5').current_region
# xlwings로 수식 -> 값
fomulaCells=r.api.SpecialCells(CellType.xlCellTypeFormulas)
xlwingsCells=sht.range(fomulaCells.GetAddress())
for cellAddress in xlwingsCells.address.split(','):
    r=sht.range(cellAddress)
    if r.columns.count!=1:
        r.value=r.value
    else:
        sr=r.options(pd.Series,index=False).value
        r.options(index=False).value=sr

#pandas 활용
# df=r.options(pd.DataFrame).value
# r.value=df
