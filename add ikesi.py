# 将小艾照片重命名为 ikesi.png 并将 Docx 文档与照片放在根目录中


from docx import Document
from docx.shared import Inches
import os, glob

run = True
dir_path = os.path.dirname(os.path.realpath(__file__))
docList = []
os.chdir(dir_path)
for file in glob.glob("*.docx"):
    i=0
    docList.insert(i, file)
    # print(file)
    i = i+1


for docName in docList:
    try:
        print(docName)
        document = Document(docName)
        document.add_picture('ikesi.png', width=Inches(0.85))
        document.save(docName)
        time.sleep(1)
    except:
        pass