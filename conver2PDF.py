# 将文件夹中的dcox文件转为pdf文件

import os, glob, time
from docx2pdf import convert

# 自动获取当前路径
dir_path = os.path.dirname(os.path.realpath(__file__))
# 创建图片文件名列表
docList = []

os.chdir(dir_path)
for file in glob.glob("*.docx"):
    i=0
    docList.insert(i, file)
    # prin t(file)
    i = i+1


for docName in docList:
    try:
        convert(docName)
        print(docName + 'convert complete')
        print('')
        # time.sleep(5)
    except:
        pass
