import os

# 实验 os.path.split 的使用方法
path = 'C:/Users/xxx/Desktop/新建文件夹\\123\\123.pdf'

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
