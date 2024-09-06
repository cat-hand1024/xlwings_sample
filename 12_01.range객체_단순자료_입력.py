import xlwings as xw

wb=xw.Book()
sht=wb.sheets(1)

sht.range('a1').value=1 # 단일자료
sht.range('a2').value=2 # 단일자료
sht.range('a3').value='합계'


