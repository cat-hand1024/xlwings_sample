import xlwings as xw
import os

filePath=r'C:\코딩학습\xlwings\예제'
fileName='009_셀To그림.xlsx'
file=os.path.join(filePath,fileName)

wb=xw.Book(file)
target_Sht=wb.sheets(1)
source_Sht=wb.sheets(2)

RangeToImage_R=source_Sht.range('b2').current_region
RangeToImage_R=RangeToImage_R.resize(RangeToImage_R.rows.count+1,None)

RangeToImage_R.copy_picture(appearance='screen',format='bitmap')

target_Sht.api.Paste(target_Sht.api.Range('d5'))

img=target_Sht.pictures[0]
img.lock_aspect_ratio=False
img.name='결제란'
img.width=target_Sht.range('d5').merge_area.width
img.height=target_Sht.range('d5').merge_area.height

target_Sht.api.Paste(target_Sht.api.Range('d5'))
