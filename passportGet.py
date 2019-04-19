#-*- coding: utf-8 -*-
#import tesserocr
#print(tesserocr.file_to_text('.\pp\LX15101140-PP'))
#import pytesseract
#from PIL import Image
import os
import re
import xlwt

def listPP():
	for root, dirs, files in os.walk('.\\passport\\'):  #root:当前目录路径; dirs:当前路径下所有子目录; files:当前路径下所有非目录子文件
		print(files)
		return files

def getInfo():
	lst_pp=listPP()
	PPdata=[['照片名','护照姓','护照名','护照号码']]
	for pic_name in lst_pp:
		try:
			result=os.popen('tesseract .\\passport\\%s rst -l eng -psm 1'%pic_name)#-psm 7
			res = result.read()  
			for line in res.splitlines():  
				print (line+'\n') 

			with open('rst.txt','r',encoding='UTF-8') as txt:
				txt=txt.read()
				Surname=re.findall(r'P<(.*?)<<',txt)[0]
				if Surname == 'IND':
					Surname=''
				givenName=re.findall(r'<<(.*?)<<<<<<<<<<<',txt)[0]
				#if '<' in givenName:
				givenName=givenName.replace('<',' ')	
				PPno=re.findall(r'<<<<<<<\n(.*?)<',txt)[0]
				PPdata.append([pic_name,Surname,givenName,PPno])
				#print ('Surname is: %s\nGiven Name is: %s\nPassport Number is: %s'%(Surname,givenName,PPno))
				print (PPdata)
		except Exception as e:
			print (pic_name,e)
			continue
	return PPdata

def wtData():
	PPdata=getInfo()
	f = xlwt.Workbook()
	sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True) #创建sheet 
	#将数据写入第 i 行，第 j 列
	i = 0
	for data in PPdata:
		for j in range(len(data)):
			sheet1.write(i,j,data[j])
		i = i + 1       
	f.save('passportInfo.xls') #保存文件

wtData()


'''
references:
https://www.cnblogs.com/Jimc/p/9772930.html
https://blog.csdn.net/nextdoor6/article/details/51283117
https://www.cnblogs.com/yizhenfeng168/p/6953330.html
def main():
	image = Image.open(".\\test\\7.jpeg")
	#image.show() #打开图片1.jpg
	text = pytesseract.image_to_string(image,lang='chi_sim') #使用简体中文解析图片
	#print(text)
	with open("output.txt", "w") as f: #将识别出来的文字存到本地
		print(text)
		f.write(str(text))
 
if __name__ == '__main__':
    main()
'''
