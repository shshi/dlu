#-*- coding: utf-8 -*-
#===========================================================
# Author：Sha0hua
# E-mail:shi.sh@foxmail.com
# Modified Date: 2019-04-17
# Version: 1.0
# Version Description: *
#===========================================================
from selenium import webdriver
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import sys
import csv
import re
import tkinter as tk
from tkinter.messagebox import *
from tkinter import *

class feecheck_scan: 
	def __init__(self,root):
		print ("initiating..." )
		self.d=webdriver.Firefox(firefoxProfile, options=options)
		self.d.set_window_size(1600, 900)
		self.d.get('http://202.203.16.20/student/')
		#time.sleep(10)
		print ("logging in...")
		try:
			self.d.find_element_by_xpath("/html/body/table/tbody/tr[2]/td[3]/form/input[1]").send_keys("jz1345")
			self.d.find_element_by_xpath("/html/body/table/tbody/tr[2]/td[3]/form/input[2]").send_keys("123456")
			self.d.find_element_by_xpath("/html/body/table/tbody/tr[2]/td[3]/form/input[3]").click()
		except Exception as e:
			print (e)
			print ("timeout")
		try:
			self.d.switch_to.frame("left")
			self.d.find_element_by_link_text("学生缴费查询").click()
			print ("successfully logged in")
		except Exception as e:
			print (e)
			print ("successfully logged in (timeout)")

		frame=Frame(root)
		frame.pack()			
		#绘制label,grid()确定行列
		Label(frame, text="学号:").grid(row=0,column=0)
		#导入输入框
		self.E1=Entry(frame,bd=2)
		#设置输入框的位置
		self.E1.grid(row=0, column=1, padx=10, pady=5)
		self.E1.bind('<Key-Return>',self.check)
		Button(frame, text="清除数据", command=self.clean).grid(row=2, column=0, sticky=W, padx=10, pady=5)
		Button(frame, text="退出", command=root.quit).grid(row=2, column=1, sticky=E, padx=10, pady=5)
		#Button(frame, text="退出后台", command=self.quitff).grid(row=2, column=1, sticky=E, padx=10, pady=5)
		self.infobox=tk.Text(root,height=50)     #这里设置文本框高，可以容纳两行
		self.infobox.pack()

	def check(self,Keyinfo):
		try:
			self.LX=self.E1.get()
			print ('	-------\n')
			self.LX.strip().strip('\n')
			self.E1.delete(0, END)
				
			self.d.switch_to.parent_frame()
			self.d.switch_to.frame("main")
			self.d.find_element_by_xpath("/html/body/form/input[2]").clear()
			self.d.find_element_by_xpath("/html/body/form/input[2]").send_keys(self.LX)
			self.d.find_element_by_xpath("/html/body/form/input[4]").click()

			table0=self.d.find_element_by_xpath("/html/body").text
			basic_info=re.findall(r'姓名：\s(.*?)\s\s费用缴纳信息',table0)[0]
			basic_info=basic_info.split(',')
			basic='\n	-------\n'+'	学号:	%s\n'%basic_info[0]+'	姓名:	%s\n'%basic_info[1]+'	学院:	%s\n'%basic_info[2]+'	班级:	%s\n'%basic_info[3]
			print (basic)
			self.infobox.insert('insert','%s\n'%basic)

			table1=self.d.find_element_by_class_name('tp')
			table_rows = table1.find_elements_by_tag_name('tr')
			due=''
			data=[]
			for tr in table_rows:      
				td =  tr.find_elements_by_tag_name('td')
				if '学号' in td[0].text:
					#print (td[0].text)
					continue
				else:
					row_val=[]
					for x in basic_info:
						x=x.strip()
						row_val.append(x)
					for cell in td:
						cell_val=cell.text
						row_val.append(cell_val)
					row_val=row_val[:4]+row_val[5:]
					if '—' not in row_val[8]:
						data.append(row_val)
						due='yes'
						due_info='	学年:	%s\n'%row_val[4]+'	项目:	%s\n'%row_val[5]+'	应收:	%s\n'%row_val[6]+'	实收:	%s\n'%row_val[7]+'	欠费:	%s\n'%row_val[8]
						print (due_info)
						self.infobox.insert('insert','%s\n'%due_info)
			#print (data)
			if due != 'yes':
				print ('该生无欠费')
				self.infobox.insert('insert','	该生无欠费!')
				#showinfo('提示','该生无欠费')            
				
		except Exception as e:
			print (e)
			LX=self.E1.get()
			self.infobox.insert('insert','	无法查到该生!%s\n'%LX)
			#showinfo('提示','无法查到该生')
			
	def clean(self):
		self.infobox.delete('0.0','end')
		
	def quitff(self,root):
		self.d.quit()
		root.quit
		
if __name__ == "__main__":
	sys.setrecursionlimit(1000000) #设置最大递归次数（若不设置，默认值为998，递归998次后将出现"maximum recursion depth exceeded"的报错）
	firefoxProfile = FirefoxProfile()
	#firefoxProfile.set_preference('permissions.default.stylesheet', 2) #禁加载CSS
	firefoxProfile.set_preference('permissions.default.image', 2) #禁加载图片
	firefoxProfile.set_preference('dom.ipc.plugins.enabled.libflashplayer.so', 'false') #禁加载Flash
	firefoxProfile.accept_untrusted_certs = True
	options = Options()
	options.add_argument('-headless') #无浏览器参数

	root=tk.Tk()
	root.title('欠费查询系统')
	#root.geometry('370x850')

	Select=tk.Label(root,text='\n请扫描学生的条形码')
	Select.pack()

	app=feecheck_scan(root)

	foot=tk.Label(root,text='\nPowered by@ShaoTech\nCopyriht © 大理大学留学生教育服务中心 版权所有\n如有问题请联系：石少华，Email: shi.sh@foxmail.com')
	foot.pack()
	root.mainloop()

	#data=[['学号','中文名','学院','班级','学年','项目名称','应收数','实收数','欠费数']]	
	print ("the end")
