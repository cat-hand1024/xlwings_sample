import xlwings as xw
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='006_01_자료형식_사용자정의.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

# 양수 -> 파란색 /음수 -> 발강색
val_01=sht.range('a2').value
val_02=sht.range('a3').value
sht.range('b2:b3').number_format='[파랑]#.##;[빨강]-#.##'
sht.range('b2').value=val_01
sht.range('b3').value=val_02

# 100이상이면 파랑 / 100이하이면 빨강
sht.range('b4:b5').number_format='[파랑][>=100];[빨강][<100]'
sht.range('b4').value=sht.range('a4').value
sht.range('b5').value=sht.range('a5').value

# 주민등록번호 '-' 적용
sht.range('b6').number_format='000000-0000000'
sht.range('b6').value=sht.range('a6').value

# 전화번호 '-' 적용
sht.range('b7:b8').number_format='[<=9999999999]"010"-00-0000;"010"-000-0000'
sht.range('b7').value=sht.range('a7').value
sht.range('b8').value=sht.range('a8').value
