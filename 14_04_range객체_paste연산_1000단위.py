import pandas as pd
import xlwings as xw
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='001_참조.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

origin_R=sht.range('b5').current_region
df=origin_R.options(pd.DataFrame,expand='table',index=False).value

# xlwings paste 연산으로 1000단위 추가
data_R=origin_R.offset(1,1).resize(origin_R.rows.count-1,origin_R.columns.count-1)

data_R.copy(destination=sht.range('n6'))

tmp_R=sht.range('n6').current_region

for colNum in range(data_R.columns.count):
    colNum=colNum-1
    tmp_R.columns(colNum).copy()
    data_R.columns(colNum).value=1000
    data_R.columns(colNum).paste(operation='multiply')
tmp_R.clear()

# pandas의 DataFrame활용
df.iloc[:,1:]=df.iloc[:,1:]*1000
sht.range('n5').options(index=False).value=df
origin_R.copy()
sht.range('n5').paste(paste='formats')

wb.api.Application.CutCopyMode=False