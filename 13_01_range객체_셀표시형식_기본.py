import pandas as pd
import xlwings as xw
import os

def formChange(routin):
    for src_num,trg_num,fm in routin:
        sr_data=r.columns(src_num).options(pd.Series,index=False).value
        r.columns(trg_num).number_format= fm
        r.columns(trg_num).options(index=False).value=sr_data

filePath=r'C:\코딩학습\xlwings\예제'
fileName='006_01_자료형식_일반.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

r=sht.range('a1').current_region
r=r.offset(2,0).resize(r.rows.count-3,None)

dataCol_num=[1,3,5,7,9,11,13]
changeCol_num=[2,4,6,8,10,12,14]
form=['0000','####','0,000.00','#,###.##','0,','#,','@"님 반갑습니다"']

routin=list(zip(dataCol_num,changeCol_num,form))

formChange(routin=routin)

# # 표시형식 : 0000
# sr_data=r.columns(1).options(pd.Series,index=False).value
# r.columns(2).number_format='0000'
# r.columns(2).options(index=False).value=sr_data
