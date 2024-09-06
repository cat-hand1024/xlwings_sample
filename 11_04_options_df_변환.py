import xlwings as xw
import os
import pandas as pd


filePath=r'C:\코딩학습\xlwings\예제'
fileName='004_영역조정.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

def change_value(row):
    if row['직위']=='사원' and row['호봉'] >=4:
        row['직위']='주임'
        row['호봉']=1
    return row


df=sht.range('b2').options(pd.DataFrame,index=False,header=1,expand='table').value
print(df)


df=df.apply(change_value,axis='columns')
sht.range('b2').options(index=False).value=df