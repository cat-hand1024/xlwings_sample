'''
총 직원수 및 직급별 인원 구하기
 range겍체에 수식 입력하여 자료 작성
'''

import xlwings as xw
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='004_영역조정.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

r=sht.range('b2').current_region
gradeR=sht.range(r.columns(3)).offset(1,0).resize(r.rows.count-1,None)

sht.range('j2').value='총 직원수'
sht.range('j3').value='부장'
sht.range('j4').value='차장'
sht.range('j5').value='과장'
sht.range('j6').value='대리'
sht.range('j7').value='주임'
sht.range('j8').value='사원'

sht.range('k2').formula='=COUNTA({})'.format(gradeR.address)
for idx,v in enumerate(['부장','차장','과장','대리','주임','사원'],start=3):
    cellName='k'+str(idx)
    sht.range(cellName).formula='=COUNTIF({},"{}")'.format(gradeR.address,v)