from docxtpl import DocxTemplate
from docx import Document
from docx2pdf import convert
import datetime
today = datetime.date.today()

name = input('输入文件名：')
student_name = input('输入学生名字：')
communication = int(input('思辨与交流能力: '))
creation = int(input('反思与创新能力：'))
co_operation = int(input('合作与互助能力：'))
solvability = int(input('问题解决能力：'))
thoughts = int(input('编程思维：'))
lecturer_comments = input('输入教师评价：')
year = input('输入年份：')
month = input('输入月份：')
day = input('输入日：')
documentName = student_name+str(year)+str(month)+str(day)+'课评报告.docx'
generatePDF = input('是否生成PDF？ YES｜NO\n')

doc = DocxTemplate(name+'.docx') #加载模板文件
document = Document(name+'.docx')


data_dic = {
    'student_name' : student_name,
    'communication' : '★'*(communication-1),
    'creation' : '★'*(creation-1),
    'co_operation' : '★'*(co_operation-1),
    'solvability' : '★'*(solvability-1),
    'thoughts' : '★'*(thoughts-1),
    'lecturer_comments' : lecturer_comments,
    'year' : year,
    'month' : month,
    'day' : day
}
doc.render(data_dic) #填充数据
doc.save(documentName) #保存目标文件1

if generatePDF == 'YES':
    convert(documentName)
else:
    pass
