# 遍历文件夹并复制特定类型(pdf)的文件并添加到压缩包中

import os, shutil, zipfile
import tkinter as tk
from tkinter import filedialog


# 使用tkinter打开窗口选择文件夹
root = tk.Tk()
root.withdraw()
original_folder_path = filedialog.askdirectory()
os.chdir(original_folder_path) # 变更程序目录

def CrossOver(dir,fl):
    for i in os.listdir(dir):  # 遍历整个文件夹
        path = os.path.join(dir, i)
        if os.path.isfile(path):  # 判断是否为一个文件，排除文件夹
            if os.path.splitext(path)[1]==".pdf":  # 判断文件扩展名是否为“.pdf”
                fl.append(path)  # 保存文件路径与文件名至fl中
        elif os.path.isdir(path):
            newdir=path
            CrossOver(newdir,fl)
    return fl

# 压缩文件
def zipdirs(dir_list):
    os.chdir(original_folder_path)
    zip_file = 'test.zip' # 压缩后的名字

    zip = zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED)
 
    for item in dir_list:
        head_tail = os.path.split(item) # 将item中的地址，拆分为地址与文件名
        zip.write(head_tail[1]) # tuple中第二个元素为文件名
    zip.close()



filelist=[]
output=CrossOver(original_folder_path, filelist)   # 执行函数，输出结果

# 复制文件到选择目录
for file in output:
    try:
        shutil.copy(file, original_folder_path)
        print(file + ' copied')
    except shutil.SameFileError:
        print("存在相同文件！")

zipdirs(output) # 压缩