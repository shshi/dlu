#-*- coding: utf-8 -*-

pool=open('pool.txt','r').read()
with open('formatch.txt','r') as fp:
	for i in fp:
		#i=i.strip('\n').upper()
		i=i.strip('\n')
		if i in pool:
			with open('matched_result.txt','a') as m:
				m.write(i+'\n')
