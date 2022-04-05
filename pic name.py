# 修改照片的名字为拍视的日期

import os
from PIL import Image

picPath = 'IMG_4093.jpg'
pict = Image.open(picPath)
exif_data = pict._getexif()
picDate = exif_data[36867]

year = picDate[0:4]
month = picDate[5:7]
day = picDate[8:10]
print(picDate)
print(type(picDate))
print(year)
print(month)
print(day)

picNameStandard = year+month+day+'_'+'1'+'.jpg'
os.rename('IMG_4093.jpg', picNameStandard)