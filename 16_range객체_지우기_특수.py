import xlwings as xw
import os
from xlwings.constants import LineStyle


filePath=r'C:\코딩학습\xlwings\예제'
fileName='011_clear.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

r=sht.range('e2').current_region
r=r.offset(2,0).resize(r.rows.count-2,None)

memoClear_R=r.columns(1)
memoClear_R.api.ClearComments()

hyperlinkClear_R=r.columns(2)
for addr in hyperlinkClear_R:
    if addr.api.Hyperlinks.Count > 0:
        # addr.api.ClearHyperlinks()
        print(addr.hyperlink)


lineClear_R=r.columns(3)
lineClear_R.api.Borders.LineStyle = LineStyle.xlLineStyleNone
