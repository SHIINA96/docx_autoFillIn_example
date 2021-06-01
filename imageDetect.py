import glob, os

# 自动获取当前路径
dir_path = os.path.dirname(os.path.realpath(__file__))

picList = []
# 从路径中查找 .jpg 文件
os.chdir(dir_path)
for file in glob.glob("*.jpg"):
    i=0
    picList.insert(i, file)
    print(file)
    i = i+1

print(picList)