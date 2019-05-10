#coding=utf-8


import os
import csv
import docx

def readdocx(file):
	# print(file)
	f = docx.Document(file)
	str = ""
	for para in f.paragraphs:
		str += para.text + " "

	idx = str.find(u"一案")
	if idx == -1:
		return ""

	return str[idx-15:idx+2].encode("utf-8")

	#print(str)



path = "C:/Users/mazy/Desktop/books/docx/" #文件夹目录
files= os.listdir(path) #得到文件夹下的所有文件名称
s = []
sf = []

for file in files: #遍历文件夹
	if not os.path.isdir(file): #判断是否是文件夹，不是文件夹才打开
		# f = open(path+"/"+file); #打开文件
		# iter_f = iter(f); #创建迭代器
		# str = ""
		# for line in iter_f: #遍历文件，一行行遍历，读取文本
		# 	str = str + line
		# if ".py" not in file:
		# 	sf.append(file)
		if ".doc" in file and "~$" not in file:
			try:
				ss = readdocx(path + file)
			except:
				print(file)
				continue
			sf.append(file)
			s.append(ss)
		# s.append(str) #每个文件的文本存到list中
          	#print(file)

with open('test2.csv','wb') as myFile:
	myWriter=csv.writer(myFile)
	# for s in sf:
	# 	myWriter.writerow([s])
	for i in range(0,len(s)):
		myWriter.writerow([s[i],sf[i]])
	# myWriter.writerow([7,sf[0]])

# print(sf)


