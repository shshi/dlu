#-*- coding: utf-8 -*-
import flask
import requests
import time
import json
app = flask.Flask(__name__)
@app.route("/")

def getList():
     
    s = requests.session()

    headers_login={"Host":"202.203.16.42",
                "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv",
                "Accept":"text/html, */*; q=0.01",
                "Accept-Language":"zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2",
                "Accept-Encoding":"gzip, deflate",
                "Content-Type":"application/x-www-form-urlencoded; charset=UTF-8",
                "X-Requested-With":"XMLHttpRequest",
                "Content-Length":"92",
                "Origin":"http",
                "Connection":"keep-alive",
                "Referer":"http"
                }

    login_data = {"username":"OTIwNzE0", "password":"MTIzNDU2", "verification":"", "token":"3578abfe-cdaf-417a-a023-6415896a2103"}
    list_data = {"type":"XSFXTWJC","xmid":"4a4b90aa73faf66a0174116ae01b0a14"}
    body_data = {"pageIndex":"0","pageSize":"40","sortField":"","sortOrder":""}
    
    '''url_code = 'http://202.203.16.42:80/nonlogin/login/captcha.htm?code=1601396928880'#验证码地址  
    temp_code = open("valcode.png","wb")
    temp_code.write(s.get(url_code).content)
    temp_code.close()
    valc = input("输入验证码：")
    login_data["verification"]=str(valc)'''

    req = s.post("http://202.203.16.42/login/Login.htm", data=login_data, headers=headers_login).cookies.get_dict()
    JSESSIONID=req['JSESSIONID']

    headers_list = {"Host":"202.203.16.42",
                "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv",
                "Accept":"text/plain, */*; q=0.01",
                "Accept-Language":"zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2",
                "Accept-Encoding":"gzip, deflate",
                "Content-Type":"application/x-www-form-urlencoded; charset=UTF-8",
                "X-Requested-With":"XMLHttpRequest",
                "Content-Length":"45",
                "Origin":"http",
                "Connection":"keep-alive",
                "Referer":"http",
                "Cookie":"username=920714; menuVisible=1; JSESSIONID=%s"%JSESSIONID
                }

    am_list = s.post("http://202.203.16.42/syt/zzglappro/queryNotExistList.htm?type=XSFXTWJC&xmid=4a4b90aa73fad84c017411601830099d", data=body_data,headers=headers_list)
    pm_list = s.post("http://202.203.16.42/syt/zzglappro/queryNotExistList.htm?type=XSFXTWJC&xmid=4a4b90aa73faf66a0174116ae01b0a14", data=body_data,headers=headers_list)

    am_data = am_list.text
    am_data=am_data.split('[')[-1].split(']')[0]
    am_data="["+am_data+"]"
    am_data = json.loads(am_data)
    amList=[i.get('xm') for i in am_data]
    if len(amList)==0:
        am_namelist=''
    else:
        am_namelist='\n上午未打卡名单：'+','.join(amList)+"; "
    print ('早上%d人未打卡 '%len(amList),am_namelist)

    pm_data = pm_list.text
    pm_data=pm_data.split('[')[-1].split(']')[0]
    pm_data="["+pm_data+"]"
    pm_data = json.loads(pm_data)
    pmList=[i.get('xm') for i in pm_data]
    if len(pmList)==0:
        pm_namelist=''
    else:
        pm_namelist='\n下午未打卡名单：'+','.join(pmList)
    print ('晚上%d人未打卡 '%len(pmList),pm_namelist)

    
if __name__ == '__main__':
    app.debug = True
    app.run().getList()
    print("All finished")    


