import xlwings as xw
import os

def multiRange_Select(*args):
    joinLst=[]

    for a in args:
        joinLst.append(a.address)

    joinRange=','.join(joinLst)

    return sht.range(joinRange)

filePath=r'C:\코딩학습\xlwings\예제'
fileName='001_참조.xlsx'
file=os.path.join(filePath,fileName)

wB=xw.Book(file)
sht=wB.sheets(1)

# 한셀 참조
cell=sht.range('b7')
cell.select()

# 연속셀 참조
r=sht.range('b5:c12')
r.select()

# 떨어진 범위 참조
# range메서드 ( 영역최저(column최저+row최저):영역최고(column최고:row최고))
r=sht.range('b5:c12','e5:f10') # column최저'b' + row최저 '5' :column최고 'f'+row최고 '12' -> range('b5 : f12')
r.select()

# 엑셀 시트 전체 행/열 선택
r=sht.cells.rows(2)
print(r.address)
r=sht.cells.columns(2)
print(r.address)

# 엑셀 시트 마지막 행/열 셀 선택
lastrow=sht.cells.last_cell.row
lastcol=sht.cells.last_cell.column
print(lastrow,lastcol)