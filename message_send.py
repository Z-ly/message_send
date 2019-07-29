# -*- coding: utf-8 -*-
"""
Created on Sun Nov 18 14:33:28 2018

@author: zly
"""

from selenium import webdriver
import xlrd

#——————打开wechat界面——————
driver=webdriver.Chrome(r'C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe')
driver.get('http://campus.info.bit.edu.cn/wechat/index')
#a = driver.find_element_by_class_name("btn").click()
driver.find_element_by_id("username").send_keys('0000000')
driver.find_element_by_id("password").send_keys('0000000')
driver.find_element_by_class_name("btn_image").click()

#—————————读取短信————————————
data = xlrd.open_workbook('ylbx.xlsx')
table = data.sheets()[0]
stu_ids = table.col_values(0)
messages = table.col_values(1)
for k in range(0,len(stu_ids)):
    print(str(int(stu_ids[k]))+messages[k])    
#    a = ['1','2','3']
#    for i in [0,1,2]:
    driver.find_element_by_id("content").send_keys(messages[k])
    driver.find_element_by_id("data").send_keys(str(int(stu_ids[k])))
    #driver.find_element_by_class_name("btn-primary").click()
