#-*- coding: utf-8 -*-
#import tesserocr
#print(tesserocr.file_to_text('.\pp\LX15101140-PP'))
import pytesseract
from PIL import Image
import os

result=os.popen("tesseract 7.jpeg 7 -l eng")#-psm 7
res = result.read()  
for line in res.splitlines():  
	print (line) 
print ('finished')

'''
references:
https://www.cnblogs.com/Jimc/p/9772930.html
https://blog.csdn.net/nextdoor6/article/details/51283117
https://www.cnblogs.com/yizhenfeng168/p/6953330.html
'''