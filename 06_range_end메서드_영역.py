import xlwings as xw
from openpyxl.utils.cell import get_column_letter
import os

def display_endCell(st_Point,arrow='left'):
    '''
     기준셀에서 상하좌우 끝영역을 출력하는 함수
    :param st_Point: 기준 셀
    :param arrow: end 방향
    :return:
    '''

    r_endAddress = st_Point.end(arrow).address
    r_endCol = st_Point.end(arrow).column
    r_endRow =st_Point.end(arrow).row

    print(r_endAddress, r_endRow, get_column_letter(r_endCol))


filePath=r'C:\코딩학습\xlwings\예제'
fileName='003_빈셀처리.xlsx'
file=os.path.join(filePath,fileName)

xb=xw.Book(file)
sht=xb.sheets(1)

# end메서드 기능 : ctr+up/down/right/left 화살표 효과

r_center=sht.range('f9')
#left
display_endCell(st_Point=r_center,arrow='left')
#right
display_endCell(st_Point=r_center,arrow='right')
#down
display_endCell(st_Point=r_center,arrow='down')
#up
display_endCell(st_Point=r_center,arrow='up')


