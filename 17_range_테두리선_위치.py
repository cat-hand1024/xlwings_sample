import xlwings as xw
from xlwings.constants import BordersIndex,LineStyle

wb=xw.Book()
sht=wb.sheets(1)

sht.range('b2').api.Borders(BordersIndex.xlEdgeTop).LineStyle=LineStyle.xlContinuous   # 위쪽 테두리
sht.range('b3').api.Borders(BordersIndex.xlEdgeLeft).LineStyle=LineStyle.xlContinuous  # 왼쪽 테두리
sht.range('b4').api.Borders(BordersIndex.xlEdgeRight).LineStyle=LineStyle.xlContinuous # 오른쪽 테두리
sht.range('b5').api.Borders(BordersIndex.xlEdgeBottom).LineStyle=LineStyle.xlContinuous # 아래쪽 테두리
sht.range('b6').api.Borders(BordersIndex.xlDiagonalUp).LineStyle=LineStyle.xlContinuous # 대각선 ( 왼쪽_하 -> 오른쪽_상 )
sht.range('b7').api.Borders(BordersIndex.xlDiagonalDown).LineStyle=LineStyle.xlContinuous # 대각선 ( 왼쪽_상 -> 오른쪽_하)
sht.range('b8:d12').api.Borders(BordersIndex.xlInsideHorizontal).LineStyle=LineStyle.xlContinuous # 테두리제외한 수평라인
sht.range('b13:d17').api.Borders(BordersIndex.xlInsideVertical).LineStyle=LineStyle.xlContinuous # 테두리제외한 수직라인
