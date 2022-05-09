# 将文件复制粘贴到同名文件夹内

import os, shutil
import tkinter as tk
from tkinter import filedialog


root = tk.Tk()
root.withdraw()
folder_path = filedialog.askdirectory()

nas_student_folder_path = "//ikesi-server/Public/8.学员课后资料/"

def traversal_folder(dir):  # 获取文件夹中的全部文件夹
    folder_list = []
    for i in os.listdir(dir):   # 遍历文件夹内的目录
        path = os.path.join(dir,i)  
        if os.path.isdir(path): # 判断目录是否为路径
            folder_list.append(path)
    return folder_list

output = traversal_folder(folder_path)  # 获取选择文件夹中的全部路径


for path in output:
    head_tail = os.path.split(path)
    folder_name = head_tail[1].split('\\')

    for file in os.listdir(path):   # 遍历文件夹中的文件
        destination_path = os.path.join(nas_student_folder_path,folder_name[0])
        if os.path.splitext(file)[1]==".jpg":
            # print(file)
            destination_path = os.path.join(destination_path, '照片')
            shutil.copy(os.path.join(path,file), destination_path)

        elif os.path.splitext(file)[1]==".MOV":
            # print(file)
            destination_path = os.path.join(destination_path, '视频')
            shutil.copy(os.path.join(path,file), destination_path)

        elif os.path.splitext(file)[1]==".pdf":
            # print(file)
            destination_path = os.path.join(destination_path, '课评报告')
            shutil.copy(os.path.join(path,file), destination_path)

        print(file + ' copied')