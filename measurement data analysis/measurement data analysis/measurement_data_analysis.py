import os
import xlwt
#定位到文件所在位置
a = os.getcwd() #获取当前目录
print (a) #打印当前目录
os.chdir('C:/Users/User/OneDrive/Miniature tripod/Test result/autocollimator drifting test') #定位到新的目录
a = os.getcwd() #获取定位之后的目录
print(a) #打印定位之后的目录
input='drifting_data'
output='output'
dataset='7'
#读取目标txt文件里的内容，并且打印出来显示
#with open('test_result_test.txt','r') as raw:
#	for line in raw:
#		print (line)

#去掉txt里面的空白行，并保存到新的文件中
with open(input+dataset+'.txt','r',encoding = 'utf-8') as fr, open(output+dataset+'.txt','w',encoding= 'utf-8') as fd:
	for text in fr.readlines():
		if text.split():
			fd.write(text)
	print('success')

#创建一个workbook对象，相当于创建一个Excel文件
book = xlwt.Workbook(encoding='utf-8',style_compression=0)
'''
Workbook类初始化时有encoding和style_compression参数
encoding:设置字符编码，一般要这样设置：w = Workbook(encoding='utf-8')，就可以在excel中输出中文了。默认是ascii。
style_compression:表示是否压缩，不常用。
'''
 
# 创建一个sheet对象，一个sheet对象对应Excel文件中的一张表格。
sheet = book.add_sheet('output', cell_overwrite_ok=True)
# 其中的Output是这张表的名字,cell_overwrite_ok，表示是否可以覆盖单元格，其实是Worksheet实例化的一个参数，默认值是False
 
# 向表中添加表题
sheet.write(0, 0, 'X')  # 其中的'0-行, 0-列'指定表中的单元，'X'是向该单元写入的内容
sheet.write(0, 1, 'Y')

#对文本内容进行多次切片得到想要的部分
n=1
with open(output+dataset+'.txt','r+') as fd:
	for text in fd.readlines():
		x=text.split(':')[2]
		y=text.split(':')[3]
		print(text.split(':'))
		print (x.split('w'))
		print (y.split('w'))
		sheet.write(n,0,x.split('w')[0])#往表格里写入X坐标 
		sheet.write(n,1,y.split('w')[0])#往表格里写入Y坐标
		n = n+1
# 最后，将以上操作保存到指定的Excel文件中
book.save(output+dataset+'.xls')  

