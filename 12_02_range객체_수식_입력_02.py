import pandas as pd
import xlwings as xw
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='005_자료입력.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

r=sht.range('b2').current_region
periodR=sht.range(r.columns(5)).offset(1,0).resize(r.rows.count-1,None)

df=r.options(pd.DataFrame,index=False,expand='table').value

df_male=df[df['성별']=='남']
df_female=df[df['성별']=='여']

sht.range('h3').value='남'
sht.range('h4').value='여'
sht.range('h5').value='평균'

sht.range('i3').value=int(df_male['근속년수'].max())
sht.range('i4').value=int(df_female['근속년수'].max())
sht.range('i5').formula='=AVERAGE({})'.format(periodR.address)