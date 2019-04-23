#-*- coding: utf-8 -*-
import pytesseract
from PIL import Image,ImageFilter
import os
import re
import time
import xlwt

def listPic():
	for root, dirs, files in os.walk('.\\passport\\'):  #root:当前目录路径; dirs:当前路径下所有子目录; files:当前路径下所有非目录子文件
		#print(files)
		return files

def ppGet(pic_name,image,info,PPdata):
	try:
		nationr=info[-2][2:5]
		nation=nationr.replace('1','I').replace('5','S').replace('8','B').replace('2','Z')
		print (nation)
		Surname=re.findall(r'%s(.*?)<<.*?'%nationr,info[-2])[0].replace('\'','').replace('’','')
		print (Surname)
		givenName=re.findall(r'<<([A-Z]*<?[A-Z]*)',info[-2])[0].replace('<',' ').replace('\'','').replace('’','')
		print (givenName)
		if '<' in info[-1][:15]:
			ppNo=re.findall(r'(^.*?)<',info[-1])[0].replace('\'','').replace('’','').replace('\t','')
			'''if ppNo[1] in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
				ppNo0=ppNo[0].replace('1','I').replace('5','S').replace('8','B').replace('2','Z')
				ppNo=ppNo0+ppNo[1]+ppNo[2:].replace('Z','2').replace('D','0').replace('O','0').replace('B','8').replace('S','5').replace('?','2').replace('I','1')'''
			print (ppNo)
			bdate=re.findall(r'^.*?<.{4}(.{6})',info[-1])[0].replace('\'','').replace('’','').replace('Z','2').replace('D','0').replace('O','0').replace('B','8').replace('S','5').replace('?','2')
			if int(bdate[0])>5:
				bdate='19'+bdate
			else:
				badate='20'+bdate
			print (bdate)
			sex=re.findall(r'^.*?<.{10}.(.)',info[-1])[0]
			if sex=='H' or sex=='M':
				sex='M'
			print (sex)
			expire=re.findall(r'^.*?<.{10}..(.{6})',info[-1])[0].replace('\'','').replace('’','').replace('Z','2').replace('D','0').replace('O','0').replace('B','8').replace('S','5').replace('?','2').replace('I','1')
			expire='20'+expire
			print (expire)
			PPdata.append([pic_name,Surname,givenName,nation,ppNo,bdate,sex,expire])
		else:
			ppNo=re.findall(r'(^.{9})',info[-1])[0].replace('\'','').replace('’','').replace('\t','')
			'''if ppNo[1] in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
				ppNo0=ppNo[0].replace('1','I').replace('5','S').replace('8','B').replace('2','Z')
				ppNo=ppNo0+ppNo[1]+ppNo[2:].replace('Z','2').replace('D','0').replace('O','0').replace('B','8').replace('S','5').replace('?','2').replace('I','1')'''
			#ppNo=ppNo[:2]+ppNo[2:].replace('Z','2').replace('D','0').replace('O','0').replace('B','8').replace('S','5').replace('?','2')
			print (ppNo)
			bdate=re.findall(r'^.{13}(.{6})',info[-1])[0].replace('\'','').replace('’','').replace('Z','2').replace('D','0').replace('O','0').replace('B','8').replace('S','5').replace('?','2').replace('I','1')
			if int(bdate[0])>5:
				bdate='19'+bdate
			else:
				badate='20'+bdate
			print (bdate)
			sex=re.findall(r'.{19}.(.)',info[-1])[0]
			if sex=='H' or sex=='M':
				sex='M'
			print (sex)
			expire=re.findall(r'.{21}(.{6})',info[-1])[0].replace('\'','').replace('’','').replace('Z','2').replace('D','0').replace('O','0').replace('B','8').replace('S','5').replace('?','2').replace('I','1')
			expire='20'+expire
			print (expire)
			PPdata.append([pic_name,Surname,givenName,nation,ppNo,bdate,sex,expire])
			#print (PPdata)
	except Exception as e:
		print (e)
		PPdata.append([pic_name])
		with open('%s.txt'%pic_name,'w',encoding='UTF-8') as er:
			er.write(str(info))
	return PPdata

def vsGet(pic_name,image,info,vsdata):
	try:
		Surname=re.findall(r'R<CHN(.*?)<<.*?',info[-2])[0].replace('\'','').replace('’','').replace('<',' ')
		print (Surname)
		givenName=re.findall(r'<<([A-Z]*<?[A-Z]*)',info[-2])[0].replace('<',' ').replace('\'','').replace('’','')
		print (givenName)
		vsNo=re.findall(r'(^.*?)<',info[-1])[0].replace('\'','').replace('’','').replace('\t','')
		print (vsNo)
		nationr=re.findall(r'%s<.(.{3}).*?'%vsNo,info[-1])[0]
		nation=nationr.replace('1','I').replace('5','S').replace('8','B').replace('2','Z')
		print (nation)
		bdate=re.findall(r'%s(.{6})'%nationr,info[-1])[0].replace('\'','').replace('’','').replace('Z','2').replace('D','0').replace('O','0').replace('B','8').replace('S','5').replace('?','2').replace('I','1')
		if int(bdate[0])>5:
			bdate='19'+bdate
		else:
			badate='20'+bdate
		print (bdate)
		sex=re.findall(r'%s.{6}.(.)'%nationr,info[-1])[0]
		if sex=='H' or sex=='M':
			sex='M'
		print (sex)
		expire=re.findall(r'%s.{6}..(.{6})'%nationr,info[-1])[0].replace('\'','').replace('’','').replace('Z','2').replace('D','0').replace('O','0').replace('B','8').replace('S','5').replace('?','2').replace('I','1')
		expire='20'+expire
		print (expire)
		vsdata.append([pic_name,Surname,givenName,nation,vsNo,bdate,sex,expire])
		#print (vsdata)
	except Exception as e:
		print (e)
		#continue
		vsdata.append([pic_name])
		with open('%s.txt'%pic_name,'w',encoding='UTF-8') as er:
			er.write(str(info))
	return vsdata

def conv(image,pic_path,pic_name):
	try:
		image=image.convert('L')
		threshold=130#130
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
		#print ('%s converted'%pic_name)
		#image=image.filter(ImageFilter.SHARPEN)
		#image.save('.\\tmp\\%s'%pic_name)
		return image
	except Exception as e:
		print (e)

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
	try:
		f.save('ppInfo.xls') #保存文件
		print ('\n信息提取已完成！')
		time.sleep(7)
	except:
		print ('警告：暂时无法写入Excel文件，请关闭名为“ppInfo.xls“的Excel文件。')
		time.sleep(10)
		wtInfo(PPdata,vsdata)

def main():
	lstPic=listPic()
	PPdata=[['照片名','护照姓','护照名','国籍','护照号码','生日','性别','护照有效期至']]
	vsdata=[['照片名','护照姓','护照名','国籍','居留许可号','生日','性别','许可有效期至']]
	for pic_name in lstPic:
		try:
			print (pic_name)
			pic_path='.\\passport\\'+pic_name
			image = Image.open(pic_path)
			image = conv(image,pic_path,pic_name)
			#txt = pytesseract.image_to_string(image,lang='eng',config='--psm 1 --oem 3 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ<0123456789')
			txt = pytesseract.image_to_string(image,lang='OcrB',config='--psm 1')
			txt=txt.upper().split('\n')
			#print (txt)
			info=[]
			for line in txt:
				line=line.strip().replace('\'','').replace('’','').replace(' ','')
				if len(re.findall(r'[^A-Z0-9<]',line)) <= 3 and len(line)>=30 and (len(re.findall(r'[0-9]',line)) > 12 or '<<' in line):
					info.append(line)
			#print (info)
			if len(info)==0:
				image=image.transpose(Image.ROTATE_180)
				txt = pytesseract.image_to_string(image,lang='OcrB',config='--psm 1 ')
				txt=txt.upper().split('\n')
				for line in txt:
					line=line.strip().replace('\'','').replace('’','').replace(' ','')
					if len(re.findall(r'[^A-Z0-9<]',line)) <= 3 and len(line)>=30 and (len(re.findall(r'[0-9]',line)) > 12 or '<<' in line):
						info.append(line)
					
			if 'R<CHN' in info[-2]:
				vsGet(pic_name,image,info,vsdata)
			else:
				ppGet(pic_name,image,info,PPdata)
		except Exception as e:
			print (e)
			PPdata.append([pic_name])
			with open('%s.txt'%pic_name,'w',encoding='UTF-8') as er:
				er.write(str(info))
	wtInfo(PPdata,vsdata)
'''
references:
https://www.cnblogs.com/Jimc/p/9772930.html
https://blog.csdn.net/nextdoor6/article/details/51283117
https://www.cnblogs.com/yizhenfeng168/p/6953330.html

降噪：
https://blog.csdn.net/t8116189520/article/details/80342512
https://blog.csdn.net/chouzhou9701/article/details/82587833

字体训练网站：
http://trainyourtesseract.com/
'''
 
if __name__ == '__main__':
    main()
