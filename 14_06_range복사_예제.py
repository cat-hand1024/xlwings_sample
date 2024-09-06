'''
엑셀 매크로&VBA 바이블 p377 예제 풀이
'''

import pandas as pd
import xlwings as xw
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='007_셀복사.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

# 복사할 영역을 DataFrame으로 전환
df=sht.range('b2').options(pd.DataFrame,index=False,expand='table').value
# boolean indexer를 통해 직위가 사원인 항목 추출
df=df[df['직위']=='대리']
# 복사할 셀에 DataFrame 입력
sht.range('I2').options(index=False).value=df

# 자료가 입력될 영역 지정
target_Range=sht.range('I2').current_region
target_Range=target_Range.offset(1,0).resize(target_Range.rows.count-1,None)

# 자료가 입력될 영역의 서식 복사
sht.range('b3').copy()
target_Range.paste(paste='formats')

# CutCopyMode 해제 : 복사 점선 없애기
wb.api.Application.CutCopyMode=False

