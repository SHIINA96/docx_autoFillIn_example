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

