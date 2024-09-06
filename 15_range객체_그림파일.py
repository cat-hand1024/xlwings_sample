import xlwings as xw
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='010_복사ToFile.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
sht=wb.sheets(1)

choice_R=sht.range('b5').current_region
choice_R.to_png(path='Range_To_Png.png')


