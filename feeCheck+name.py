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
 
def login():
    d.get('http://202.203.16.20/student/')
    #time.sleep(10)
    print ("logging in...")
    try:
        d.find_element_by_xpath("/html/body/table/tbody/tr[2]/td[3]/form/input[1]").send_keys("jz1345")
        d.find_element_by_xpath("/html/body/table/tbody/tr[2]/td[3]/form/input[2]").send_keys("123456")
        d.find_element_by_xpath("/html/body/table/tbody/tr[2]/td[3]/form/input[3]").click()
    except Exception as e:
        print (e)
        print ("timeout")
    try:
        d.switch_to.frame("left")
        d.find_element_by_link_text("学生缴费查询").click()
        print ("successfully logged in")
    except Exception as e:
        print (e)
        print ("successfully logged in (timeout)")

def check(i):
    try:
        d.switch_to.parent_frame()
        d.switch_to.frame("main")
        d.find_element_by_xpath("/html/body/form/input[2]").clear()
        d.find_element_by_xpath("/html/body/form/input[2]").send_keys(i)
        d.find_element_by_xpath("/html/body/form/input[4]").click()

        table0=d.find_element_by_xpath("/html/body").text
        basic_info=re.findall(r'姓名：\s(.*?)\s\s费用缴纳信息',table0)[0]
        basic_info=basic_info.split(',')

        table1=d.find_element_by_class_name('tp')
        table_rows = table1.find_elements_by_tag_name('tr')
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
    except Exception as e:
        #print (e)
        print ('未查到该生：%s'%i)
        with open('error.txt','a') as er:
            er.write(i+'\n')

def loop():
    with open('list.txt','r') as fp:
        for i in fp:
            i=i.strip().strip('\n')
            print (i)
            check(i)


def writecsv():	
    with open('fee_result.csv', 'a', newline='') as csvfile:
        writer  = csv.writer(csvfile)
        for row in data:
            writer.writerow(row)

if __name__ == "__main__":
    sys.setrecursionlimit(1000000) #设置最大递归次数（若不设置，默认值为998，递归998次后将出现"maximum recursion depth exceeded"的报错）
    firefoxProfile = FirefoxProfile()
    #firefoxProfile.set_preference('permissions.default.stylesheet', 2) #禁加载CSS
    firefoxProfile.set_preference('permissions.default.image', 2) #禁加载图片
    firefoxProfile.set_preference('dom.ipc.plugins.enabled.libflashplayer.so', 'false') #禁加载Flash
    firefoxProfile.accept_untrusted_certs = True
    options = Options()
    #options.add_argument('-headless') #无浏览器参数
    d=webdriver.Firefox(firefoxProfile, options=options)
    d.set_window_size(1600, 900)
    print ("initiating..." )
    data=[['学号','中文名','学院','班级','学年','项目名称','应收数','实收数','欠费数']]	
    login()
    print ('checking...\n')
    loop()
    writecsv()
    print ("the end")
    time.sleep(7)
    d.quit()
