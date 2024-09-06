import pandas as pd
import xlwings as xw
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='004_영역조정.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

#연속된 두컬럼 시리즈만들기
#사번과 이름 추출
r=sht.range('b2').current_region
srR=sht.range(r.columns(1),r.columns(2))
sr=srR.options(pd.Series).value

#사번과 이름을 i2cell에 입력
sht.range('i2').value=sr

#이름만 L2 셀에 입력
sht.range('L2').options(index=False).value=sr

# 한개의 컬럼 시리즈만들기
srR=sht.range(r.columns(6))
sr=srR.options(pd.Series,index=False).value # index는 자동생성되도록 index=False를 적용
sht.range('M2').options(index=False).value=sr