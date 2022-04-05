# 将照片转为默认的角度

from PIL import Image, ExifTags

numbers = input('numbers of pictures: ')

for i in range(int(numbers)):
    picName = input('name of pictures: ')
    try:
        image=Image.open(picName)

        for orientation in ExifTags.TAGS.keys():
            if ExifTags.TAGS[orientation]=='Orientation':
                break
        
        exif = image._getexif()

        try:
            if exif[orientation] == 3:
                image=image.rotate(180, expand=True)
            elif exif[orientation] == 6:
                image=image.rotate(270, expand=True)
            elif exif[orientation] == 8:
                image=image.rotate(90, expand=True)
            print('process done!')
        except:
            print('no need to process!')

        image.save(picName)
        image.close()
    except (AttributeError, KeyError, IndexError):
        # cases: image don't have getexif
        print('no need to process!')