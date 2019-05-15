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
 
def login():
    d.get('http://dali.at0086.cn/Login.aspx')
    #time.sleep(10)
    print ("logging in...")
    try:
        d.find_element_by_id("txtLoginName").send_keys("dali_wayne")
        d.find_element_by_id("txtPassWord").send_keys("123456")
        yzm=input('请输入验证码：')
        d.find_element_by_id("txtCode").send_keys(yzm)
        d.find_element_by_id("btnLogin").click()
        print ("登录成功")
        time.sleep(3)
        ul=d.find_element_by_id('menu-top')
        ul.find_elements_by_xpath('li')[2].click()
        d.switch_to.frame(1)
    except Exception as e:
        print ("登录失败")
        print (e)
       

def check(i):
    try:
        d.find_element_by_id('ContentPlaceHolder1_txtStuNo').clear()
        d.find_element_by_id('ContentPlaceHolder1_txtStuNo').send_keys(i)
        time.sleep(1)
        d.find_element_by_id('btn_search').click()
        source=d.page_source
        with open('source.txt','w',encoding='UTF-8') as res:
            res.write(source)
        table=d.find_element_by_id('list')
        table_rows = table.find_elements_by_tag_name('td')
        row_val=[]
        for x in table_rows:      
            #td =  tr.find_elements_by_tag_name('td')
            x=x.text
            if '\n' in x:
                x=x.split('\n')
                for xx in x:
                    row_val.append(xx)
            else:
                row_val.append(x)
        print (row_val)
        if 'Data' in row_val[0]:
            print ('未查到该生：%s'%i)
            with open('error.txt','a') as er:
                er.write(i+'\n')
            row_val=[]
        else:
            data.append(row_val)
            #print (data)
    except Exception as e:
        print (e)

def loop():
    with open('list.txt','r') as fp:
        for i in fp:
            #print (i)
            if len(i)<2:
                continue
            else:
                i=i.strip().strip('\n')
                check(i)

def writecsv():	
    with open('list_result.csv', 'a', newline='') as csvfile:
        writer  = csv.writer(csvfile)
        for row in data:
            writer.writerow(row)

if __name__ == "__main__":
    sys.setrecursionlimit(1000000) #设置最大递归次数（若不设置，默认值为998，递归998次后将出现"maximum recursion depth exceeded"的报错）
    firefoxProfile = FirefoxProfile()
    #firefoxProfile.set_preference('permissions.default.stylesheet', 2) #禁加载CSS
    #firefoxProfile.set_preference('permissions.default.image', 2) #禁加载图片
    firefoxProfile.set_preference('dom.ipc.plugins.enabled.libflashplayer.so', 'false') #禁加载Flash
    firefoxProfile.accept_untrusted_certs = True
    options = Options()
    #options.add_argument('-headless') #无浏览器参数
    d=webdriver.Firefox(firefoxProfile, options=options)
    d.set_window_size(1600, 900)
    print ("initiating..." )
    data=[['序号','申请编号','学号','护照姓名','中文姓名','护照号','国籍','专业','学院','经费来源','学生类别','学习期限','修改日期','修改时间','状态','操作']]	
    login()
    print ('checking...\n')
    loop()
    writecsv()
    print ("the end")
    time.sleep(30)
    d.quit()
