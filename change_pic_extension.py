# 将NAS中照片的后缀名统一修改为'.jpg'
import os

pic_type = ['.jpg','.JPG','.JPEG','.jpeg']
path = "//ikesi-server/图图/传输"

for file in os.listdir(path):
    pic_path = os.path.join(path,file)
    # print(os.path.splitext(pic_path)[1])
    if os.path.splitext(pic_path)[1] in pic_type:
        # print(os.path.splitext(pic_path))
        new_name = os.path.splitext(pic_path)[0] + '.jpg'
        os.replace(pic_path,new_name)