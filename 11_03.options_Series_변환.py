import xlwings as xw
import os
import pandas as pd


filePath=r'C:\코딩학습\xlwings\예제'
fileName='004_영역조정.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

#사번과 이름 추출
# !!!주의 연속자료만 적용
r=sht.range('b2').current_region
srR=sht.range(r.columns(1),r.columns(2))
srR.select()
sr_staff=srR.options(pd.Series).value
print(sr_staff)
