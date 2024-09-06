import xlwings as xw
import os

def multiRange_Select(*args):
    ''' 영역 Union : Ctr+Select'''
    joinLst=[x.address for x in args]

    joinRange=','.join(joinLst)
    return sht.range(joinRange)

filePath=r'C:\코딩학습\xlwings\예제'
fileName='001_참조.xlsx'
file=os.path.join(filePath,fileName)

wB=xw.Book(file)
sht=wB.sheets(1)

# current_region :  선택된 셀이 속한 연속된 영역 모두를 선택함.
r=sht.range('b5').current_region
r.select()

# 선택 영역 행/열 전체 선택하기
choice_row=r.rows(1)
choice_row.select()
choice_col=r.columns(1)
choice_col.select()

# expand메서드
r=sht.range('b6').expand('right') # b5에서 오른쪽 연속영역 ( Ctr+오른쪽화살표 )
r.select()
r=sht.range('c5').expand('down') # b5에서 아래쪽 연속영역 ( Ctr+ 아래쪽화살표 )
r.select()
r=sht.range('b7').expand('table') # Ctr+오른쪽화살표 -> 아래쪽화살표
r.select()
#
# 셀이름으로 참조
r=sht.range('start').current_region
print(r.address)
#
# 멀티 영역 선택
multiR=multiRange_Select(r.columns(1),r.columns(3),r.columns(5))
multiR.select()

# 엑셀 시트 전체 행/열 선택
r=sht.cells.rows(2)
print(r.address)
r=sht.cells.columns(2)
print(r.address)