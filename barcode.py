#-*- coding: utf-8 -*-
#===========================================================
# Author：Sha0hua
# E-mail:shi.sh@foxmail.com
# Modified Date: 2019-04-17
# Version: 1.0
# Version Description: *
#===========================================================
import code128
from PIL import Image,ImageDraw,ImageFont
import re

with open('roll.txt','r') as roll:
	for i in roll:
		i=i.strip().strip('\n')
		i_lst=list(i)
		fstN=re.findall(r'\d',i)[0]
		pstn=i_lst.index(fstN)
		i_lst.insert(pstn,'1')
		i_new=''.join(i_lst)
		#print (i_new)
		
		#code128.image("LX14164152").save("LX14164152.jpg")
		ttfont = ImageFont.truetype("C:\Windows\Fonts\Arial.ttf",20)
		bar = code128.image(i_new)
		new_im = Image.new("RGB", (450, 200),(255,255,255))#396,200
		new_im.paste(bar,(0,33))#30
		draw = ImageDraw.Draw(new_im)
		draw.text((130,150),'%s'%i,fill=(0),font=ttfont)
		#new_im.show()
		new_im.save('%s.jpg'%i)
