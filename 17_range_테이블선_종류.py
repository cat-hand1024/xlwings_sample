import xlwings as xw
from xlwings.constants import LineStyle,BorderWeight
from xlwings.utils import rgb_to_int


wb=xw.Book()
sht=wb.sheets(1)

#모양지정
sht.range('b2').api.Borders.LineStyle=LineStyle.xlDash
sht.range('b4').api.Borders.LineStyle=LineStyle.xlContinuous
sht.range('b6').api.Borders.LineStyle=LineStyle.xlDot
sht.range('b8').api.Borders.LineStyle=LineStyle.xlDouble
sht.range('b10').api.Borders.LineStyle=LineStyle.xlDashDot
sht.range('b12').api.Borders.LineStyle=LineStyle.xlDashDotDot
sht.range('b14').api.Borders.LineStyle=LineStyle.xlSlantDashDot

for r in range(2,9,2):
    sht.range(r,4).api.Borders.LineStyle=LineStyle.xlContinuous

#굵기지정
sht.range('d2').api.Borders.Weight=BorderWeight.xlThin
sht.range('d4').api.Borders.Weight=BorderWeight.xlThick
sht.range('d6').api.Borders.Weight=BorderWeight.xlMedium
sht.range('d8').api.Borders.Weight=BorderWeight.xlHairline

# 색상지정
sht.range('d2').api.Borders.Color=int('120907',16)
sht.range('d4').api.Borders.ColorIndex=50
sht.range('d6').api.Borders.Color=rgb_to_int((205,246,187))
