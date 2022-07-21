# 将文件复制粘贴到同名文件夹内

import os, shutil
import tkinter as tk
from tkinter import filedialog


root = tk.Tk()
root.withdraw()
folder_path = filedialog.askdirectory()

nas_student_folder_path = "//ikesi-server/Public/8.学员课后资料/"
# nas_student_folder_path = "C:/Users/巴图/Desktop/新建文件夹 - 副本"

def traversal_folder(dir):  # 获取文件夹中的全部文件夹
    folder_list = []
    for i in os.listdir(dir):   # 遍历文件夹内的目录
        path = os.path.join(dir,i)  
        if os.path.isdir(path): # 判断目录是否为路径
            folder_list.append(path)
    return folder_list

output = traversal_folder(folder_path)  # 获取选择文件夹中的全部路径


def exist_and_rename(destination_path, ori_folder_path, file_name): # 判断文件是否已经存在并重新命名
    if os.path.exists(os.path.join(destination_path, file_name)):
        if file_name.endswith('.jpg'):   # 判断后缀名
            new_name=file_name.replace('.jpg','_1.jpg')   # 在后缀名前添加_1

        if file_name.endswith('.MOV'):   # 判断后缀名
            new_name=file_name.replace('.MOV','_1.MOV')   # 在后缀名前添加_1

        if file_name.endswith('.pdf'):   # 判断后缀名
            new_name=file_name.replace('.pdf','_1.pdf')   # 在后缀名前添加_1

        else:
            new_name = file_name

        new_file_path = os.path.join(ori_folder_path,new_name)
        os.rename(os.path.join(ori_folder_path,file_name),os.path.join(ori_folder_path,new_name))
        print(file_name + ' 重命名为 ' + new_name)
        return new_file_path
    else:
        return os.path.join(ori_folder_path, file_name)


for path in output:
    head_tail = os.path.split(path)
    folder_name = head_tail[1].split('\\')

    for file in os.listdir(path):   # 遍历文件夹中的文件
        if os.path.splitext(file)[1]!=".docx" and os.path.splitext(file)[1]!=".docm":  # 排除word文件
            destination_path = os.path.join(nas_student_folder_path,folder_name[0])
            file_path = os.path.join(path,file)
            if os.path.splitext(file)[1]==".jpg":
                destination_path = os.path.join(destination_path, '照片')
                
            elif os.path.splitext(file)[1]==".MOV":
                destination_path = os.path.join(destination_path, '视频')

            elif os.path.splitext(file)[1]==".pdf":
                destination_path = os.path.join(destination_path, '课评报告')

            else:
                pass

            new_file = exist_and_rename(destination_path, path, file)

            try:
                shutil.copy(new_file, destination_path)
                print(folder_name[0] + ' ' + file + ' copied')
            except FileNotFoundError:
                print('学生名错误，', folder_name[0], '不存在')

input('Press Enter to exit…')