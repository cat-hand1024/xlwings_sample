import pandas as pd
import xlwings as xw
import os


filePath=r'C:\코딩학습\xlwings\예제'
fileName='004_영역조정.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

sr=sht.range('d2').options(pd.Series,expand='down',index=False).value
totalCount_Staff=sr.count()
sr_gradeCont=sr.value_counts()
input_dict=sr_gradeCont.to_dict()

sht.range('I2').value='총 사원수'
sht.range('J2').value=totalCount_Staff

sht.range('I3').value=input_dict


