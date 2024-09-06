import pandas as pd
import xlwings as xw
import os


filePath=r'C:\코딩학습\xlwings\예제'
fileName='015_문자숫자To숫자형.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

origin_R=sht.range('b2').current_region
totalSum_R=origin_R.columns(origin_R.columns.count).offset(1,0).resize(origin_R.rows.count-1,None)

# xlwings 활용
# sht.range('g3').formula='=average({})'.format(totalSum_R.address)
#
# totalSum_R.copy(destination=sht.range('h3'))
# tmp_R=sht.range('h3').expand('down')
# tmp_R.copy()
# totalSum_R.value=1
# totalSum_R.paste(operation='multiply')
#
# tmp_R.clear()

# pandas 활용
sr=totalSum_R.options(pd.Series,expand='down',index=False).value
sr=sr.astype('int32')
totalSum_R.options(index=False).value=sr



