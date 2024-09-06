import xlwings as xw
import os

def change_form(r):
    for f,v in r:
        target_add=sht.range(v.row,v.column+1)
        target_add.number_format=f.value
        target_add.value=v.value


filePath=r'C:\코딩학습\xlwings\예제'
fileName='006_01_자료형식_날짜.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

r=sht.range('a1').current_region.offset(1,0)
fm=r.columns(1)
data=r.columns(2)
join=zip(fm,data)

change_form(r=join)

# srcVal=sht.range('b2').value
# sht.range('c2').number_format='yy'
# sht.range('c2').value=srcVal
