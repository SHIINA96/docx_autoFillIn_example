import os, re, string

# 实验 os.path.split 的使用方法
path = 'C:/Users/xxx/Desktop/新建文件夹/12/aa123.pdf'

head_tail = os.path.split(path)

print("Head of '% s:'" % path, head_tail[0]) 
print("Tail of '% s:'" % path, head_tail[1], "\n") 
print(type(head_tail))
""""
输出结果
Head of 'C:/Users/xxx/Desktop/新建文件夹\123\123.pdf:' C:/Users/xx/Desktop/新建文件夹\123
Tail of 'C:/Users/xxx/Desktop/新建文件夹\123\123.pdf:' 123.pdf 
<class 'tuple'>
"""

# 实验使用 os.path.splitext 进一步分离地址下的文件名与后缀名
file_type = os.path.splitext(head_tail[1])
print("Head of '% s:'" % path, file_type[0]) 
print("Tail of '% s:'" % path, file_type[1], "\n") 
print(type(file_type))
""""
输出结果
Head of 'C:/Users/xxx/Desktop/新建文件夹\123\123.pdf:' 123
Tail of 'C:/Users/xxx/Desktop/新建文件夹\123\123.pdf:' .pdf
<class 'tuple'>
"""

# 实验使用正则判断文件名中是否包含数字
pattern = re.compile('[0-9]+')
match = pattern.findall(head_tail[1])
if match:
    print ('contains digital')
else:
    print ('not contains digital')
""""
输出结果
contains digital
"""


# 实验使用 isdigit 判断字符串是否以数字开头
if head_tail[1].isdigit():
    print('start with digit')
else:
    print('not start with digit')
""""
输出结果
not start with digit
"""


# 将文件地址拆分并获取文件夹名称
head_tail = os.path.split(path)
file_dir = head_tail[0].split('/')
print(file_dir)
print(file_dir[-1])
""""
['C:', 'Users', 'xxx', 'Desktop', '新建文件夹', '12']
12
"""



# 将文件复制到本地网络的NAS中
import os, shutil
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()
folder_path = filedialog.askdirectory()

nas_student_folder_path = "//ikesi-server/Public/8.学员课后资料/"   # 此处为NAS文件夹的路径，注意使用'/'为路径中的分隔符，否则会报错

def traversal_folder(dir):  # 获取文件夹中的全部文件夹
    folder_list = []
    for i in os.listdir(dir):   # 遍历文件夹内的目录
        path = os.path.join(dir,i)  
        if os.path.isdir(path): # 判断目录是否为路径
            folder_list.append(path)
    return folder_list

output = traversal_folder(folder_path)  # 获取选择文件夹中的全部路径

folder1_path = output[0]
print(folder1_path)
head_tail = os.path.split(folder1_path)
folder1_name = head_tail[1].split('\\')
print(folder1_name[0])


for file in os.listdir(folder1_path):
    destination_path = os.path.join(nas_student_folder_path,folder1_name[0])
    if os.path.splitext(file)[1]==".jpg":
        # print(file)
        destination_path = os.path.join(destination_path, '照片')   # 在学生姓名后再添加一级目录
        shutil.copy(os.path.join(folder1_path,file), destination_path)


# 修改文件名
path = "C:/Users/巴图/Desktop/新建文件夹/高梓函/20220508_1.jpg"
head_tail = os.path.split(path)
if head_tail[1].endswith('.jpg'):   #当文件名以.txt后缀结尾时
    old_name = head_tail[1]
    new_name=old_name.replace('.jpg','_1.jpg')   #将原来名字里的‘.txt’替换为‘-test.txt’,相当于加个后缀了
    os.rename(os.path.join(head_tail[0],old_name),os.path.join(head_tail[0],new_name))

# 文件是否存在
os.path.exists("C:/Users/巴图/Desktop/新建文件夹/高梓函/20220508_1.jpg")