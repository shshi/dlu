#-*- coding: utf-8 -*-
#===========================================================
# Author: Sha0hua
# E-mail: shi.sh@foxmail.com
# Modified Date: 2019-06-21
#===========================================================
import code128
from PIL import Image,ImageDraw,ImageFont
import re

with open('roll.txt','r') as roll:
	for i in roll:
		i=i.strip().strip('\n')
		
		'''i_lst=list(i)
		fstN=re.findall(r'\d',i)[0]
		pstn=i_lst.index(fstN)
		i_lst.insert(pstn,'1')
		i_new=''.join(i_lst)
		print (i_new)'''
		
		#code128.image("LX14164152").save("LX14164152.jpg")
		ttfont = ImageFont.truetype("C:\Windows\Fonts\Arial.ttf",20)
		bar = code128.image(i)
		org_width, org_height = bar.size
		print (org_width, ' ', org_height)
		
		new_im = Image.new("RGB", (org_width, org_height+90),"white")#396,200
		new_im.paste(bar,(0,35))#30
		draw = ImageDraw.Draw(new_im)
		w, h = draw.textsize(i,font=ttfont)
		draw.text((org_width/2-w/2,org_height+50),i,fill="black",font=ttfont)
		#new_im.show()
		new_im.save('%s.jpg'%i)
