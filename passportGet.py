#-*- coding: utf-8 -*-
#import tesserocr
#print(tesserocr.file_to_text('.\pp\LX15101140-PP'))
#import pytesseract
from PIL import Image,ImageFilter
import os
import re
import xlwt

def listPic():
	for root, dirs, files in os.walk('.\\passport\\'):  #root:当前目录路径; dirs:当前路径下所有子目录; files:当前路径下所有非目录子文件
		#print(files)
		return files

def vsGet(pic_name,txt,vsdata):
	try:
		Surname=re.findall(r'R<CHN(.*?)<<',txt)[0].replace('\'','').replace('’','').replace(' ','')
		givenName=re.findall(r'R<CHN.*?<<(.*?)<<<',txt)[0].replace(' ','').replace('<',' ').replace('\'','').replace('’','')
		if givenName=='' and '<' in Surname:
			Surname=Surname.replace('<',' ')
			givenName=Surname
			Surname=''
		vs_num=re.findall(r'[\s]+(.*?)<.*?<<',txt)
		#print (vs_num)
		vs_num=[i for i in vs_num if i !='']
		vs_num=vs_num[-1].replace('\'','').replace('’','').replace(' ','').replace('z','2').replace('Z','2').replace('D','0').replace('O','0').replace('o','0')
		mix=re.findall(r'\n.*?<(.*?)<<',txt)[1].replace('\'','').replace('’','').replace(' ','').replace('z','2').replace('Z','2').replace('D','0').replace('O','0').replace('o','0')
		bdate=mix[4:10]
		if int(bdate[0])>5:
			bdate='19'+bdate
		else:
			badate='20'+bdate
		sex=mix[11:12]
		if sex=='H' or sex=='M':
			sex='男'
		else:
			sex='女'
		expire='20'+mix[12:18]
		vsdata.append([pic_name,Surname,givenName,vs_num,bdate,sex,expire])
		#print (vsdata)
	except Exception as e:
		print (pic_name,e)
		#continue
		vsdata.append([pic_name])
	return vsdata
		
def ppGet(pic_name,txt,PPdata):
	try:
		Surname=re.findall(r'P<(.*?)<<',txt)[0].replace('\'','').replace('’','').replace(' ','')
		#print (Surname)
		if len(Surname)==3:
			Nation=Surname
			Surname=''
		else:
			Nation=Surname[0:3]
			Surname=Surname[3:]
		givenName=re.findall(r'<<(.*?)<<<<<<<<<<<',txt)[0].replace(' ','').replace('<',' ').replace('\'','').replace('’','')
		PPno=re.findall(r'<<<<<<<[\s]+(.*?)<',txt)[0].replace('\'','').replace('’','').replace(' ','').replace('\n','').replace('\t','')
		#print (txt)
		#print ('ppno: %s'%PPno)
		mix=re.findall(r'<(.*?)<<<<',txt)
		mix=[i for i in mix if i !='']
		mix=mix[-1].replace('\'','').replace('’','').replace(' ','').replace('z','2').replace('Z','2').replace('D','0').replace('O','0').replace('o','0')
		#print (mix)
		bdate=mix[4:10]
		if int(bdate[0])>5:
			bdate='19'+bdate
		else:
			badate='20'+bdate
		sex=mix[11:12]
		if sex=='H' or sex=='M':
			sex='男'
		else:
			sex='女'
		expire='20'+mix[12:18]
		PPdata.append([pic_name,Surname,givenName,Nation,PPno,bdate,sex,expire])
		#print (PPdata)
	except Exception as e:
		print (pic_name,e)
		PPdata.append([pic_name])
		#continue
	return PPdata

def wtInfo(PPdata,vsdata):
	f = xlwt.Workbook()
	
	sheet1 = f.add_sheet(u'pp',cell_overwrite_ok=True) #创建sheet 
	#将数据写入第 i 行，第 j 列
	i = 0
	for data in PPdata:
		for j in range(len(data)):
			sheet1.write(i,j,data[j])
		i = i + 1
		
	sheet2 = f.add_sheet(u'vs',cell_overwrite_ok=True) #创建sheet 
	#将数据写入第 i 行，第 j 列
	i = 0
	for data in vsdata:
		for j in range(len(data)):
			sheet2.write(i,j,data[j])
		i = i + 1
		
	f.save('passportInfo.xls') #保存文件
	
def split():
	lstPic=listPic()
	vsdata=[['照片名','护照姓','护照名','居留许可号','生日','性别','许可有效期至']]
	PPdata=[['照片名','护照姓','护照名','国籍','护照号码','生日','性别','护照有效期至']]
	for pic_name in lstPic:
		try:
			print (pic_name)
			result=os.popen('tesseract .\\tmp\\%s rst -l eng --psm 1'%pic_name)#-psm 7
			res = result.read()  
			for line in res.splitlines():  
				print (line+'\n') 
			with open('rst.txt','r',encoding='UTF-8') as txt:
				txt=txt.read()
				if 'R<CHN' in txt:
					vsGet(pic_name,txt,vsdata)
				else:
					ppGet(pic_name,txt,PPdata)
		except Exception as e:
			print (pic_name,e)
	wtInfo(PPdata,vsdata)

def conv():
	lstPic=listPic()
	for pic_name in lstPic:
		path='.\\passport\\'+pic_name
		image=Image.open(path)
		#image.filter(ImageFilter.SMOOTH).save('.\\tmp\\11-%s'%pic_name,quality=95)
		image=image.convert('L')
		threshold=110#130
		table=[]
		for i in range(256):
			if i<threshold:
				table.append(0)
			else:
				table.append(1)
		image=image.point(table,'1')
		#image.show()
		

		data = image.getdata()
		w,h = image.size
		black_point = 0
		 
		for x in range(1,w-1):
			for y in range(1,h-1):
				mid_pixel = data[w*y+x] # 中央像素点像素值
				if mid_pixel <50: # 找出上下左右四个方向像素点像素值
					top_pixel = data[w*(y-1)+x]
					left_pixel = data[w*y+(x-1)]
					down_pixel = data[w*(y+1)+x]
					right_pixel = data[w*y+(x+1)]
		 
					# 判断上下左右的黑色像素点总个数
					if top_pixel <10:#10
						black_point += 1
					if left_pixel <10:
						black_point += 1
					if down_pixel <10:
						black_point += 1
					if right_pixel <10:
						black_point += 1
					if black_point <1:
						image.putpixel((x,y),255)
					#print('blackpoint: %d'%black_point)
					black_point = 0
		image.filter(ImageFilter.SMOOTH)
		image.save('.\\tmp\\%s'%pic_name)
		print ('%s converted'%pic_name)

conv()
split()


'''
references:
https://www.cnblogs.com/Jimc/p/9772930.html
https://blog.csdn.net/nextdoor6/article/details/51283117
https://www.cnblogs.com/yizhenfeng168/p/6953330.html
降噪：
https://blog.csdn.net/t8116189520/article/details/80342512
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
