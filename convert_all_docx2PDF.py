# 将文件夹以及子文件夹中的dcox文件转为pdf文件

import os, glob, time
import tkinter as tk
from tkinter import filedialog
from docx2pdf import convert

# 使用tkinter打开窗口选择文件夹
root = tk.Tk()
root.withdraw()
folder_path = filedialog.askdirectory()

def CrossOver(dir,fl):
    for i in os.listdir(dir):  # 遍历整个文件夹
        path = os.path.join(dir, i)
        if os.path.isfile(path):  # 判断是否为一个文件，排除文件夹
            if os.path.splitext(path)[1]==".docx":  # 判断文件扩展名是否为“.docx”
                try:
                    convert(path)
                    time.sleep(1)
                    print(path + ' convert complete')
                except:
                    pass
                fl.append(i)
        elif os.path.isdir(path):
            newdir=path
            CrossOver(newdir,fl)
    return fl


directory=folder_path  #文件夹名称
filelist=[]
output=CrossOver(directory,filelist)   # 执行函数，输出结果


input('Press Enter to exit…')