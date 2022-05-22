# 将文件夹中的照片复制到NAS中的指定文件夹

import os, shutil
import tkinter as tk
from tkinter import filedialog

# nas_student_folder_path = "//ikesi-server/Public/14.校外课程影视资料/2022 春 博苑澳森/启蒙2班/"
nas_student_folder_path = "//ikesi-server/Public/14.校外课程影视资料/2022 春 南开小学/三年四班/"

root = tk.Tk()
root.withdraw()
folder_path = filedialog.askdirectory()

def traversal_folder(dir):
    pic_list = []
    for i in os.listdir(dir):
        path = os.path.join(dir,i)  
        if os.path.isfile(path):
            pic_list.append(path)
    return pic_list

output = traversal_folder(folder_path)


for path in output:
    shutil.copy(path, nas_student_folder_path)
    print(path + ' copied')