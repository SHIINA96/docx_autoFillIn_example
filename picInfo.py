from PIL import Image

picPath = 'WechatIMG417.jpeg'
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