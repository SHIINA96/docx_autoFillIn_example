# 将文件夹以及子文件夹中的excel文件转为pdf文件

import os, time
import tkinter as tk
from tkinter import filedialog
from win32com import client


# 使用tkinter打开窗口选择文件夹
root = tk.Tk()
root.withdraw()
folder_path = filedialog.askdirectory()

def convert(origional_file_path):
    file_path = os.path.splitext(origional_file_path)[0]
    export_file_path = file_path + '.pdf'

    excel = client.Dispatch("Excel.Application")
    sheets = excel.Workbooks.Open(origional_file_path)      # Read Excel File
    work_sheets = sheets.Worksheets[0]
    work_sheets.ExportAsFixedFormat(0, export_file_path)        # Convert into PDF File

    excel.Quit()

def CrossOver(dir,fl):
    for i in os.listdir(dir):  # 遍历整个文件夹
        path = os.path.join(dir, i)
        if os.path.isfile(path):  # 判断是否为一个文件，排除文件夹
            if os.path.splitext(path)[1]==".xlsx":  # 判断文件扩展名是否为“.xlsx”
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