import xlwings as xw
import os


filePath=r'C:\코딩학습\xlwings\예제'
fileName='000_자료입력.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

curRange=sht.range('a1').current_region

# 1. 현재영역의 호봉 컬럼 선택 -> columns(5)
# 2. 선택위치 조정 ( row방향으로 +1 ) -> offset(1,0)
# 3. 영역 크기 조정 ( 현재영역 row갯수 -1, None)
numberTypeChange=curRange.columns(5).offset(1,0).resize(curRange.rows.count-1,None) # 호봉 컬럼의 Data영역만을 선택
originValue=numberTypeChange.value
changeValue=numberTypeChange.options(numbers=int).value
print(originValue) # [13.0, 9.0, 4.0, 0.0, 3.0, 4.0, 1.0, 0.0, 1.0]
print(changeValue) # [13, 9, 4, 0, 3, 4, 1, 0, 1]
