import xlwings as xw
import os


filePath=r'C:\코딩학습\xlwings\예제'
fileName='000_자료입력.xlsx'
file=os.path.join(filePath,fileName)


wb=xw.Book(file)
sht=wb.sheets(1)

r=sht.range('a1').current_region
print(r.address)

# offset(row,column) : table영역을 지정된 만큼 크기 변동없이 영역이동한다.

r=r.offset(1,0) # row방향으로 +1
print(r.address)
r=r.offset(0,2)
print(r.address)

# 응용 1 : 행추가 위치 찾기
# row 마지막셀찾기
sheetLastCellRow=sht.range(sht.cells.last_cell.row,2)
insertRowCell=sheetLastCellRow.end('up').offset(1,0) # 기존자료에서 row방향으로 +1
print(insertRowCell.address)

# 응용 2 : 영추가 위치 찾기
# column 마지막셀찾기
sheetLastCellCol=sht.range(1,sht.cells.last_cell.column)
insertColCell=sheetLastCellCol.end('left').offset(0,1)
print(insertColCell.address)

