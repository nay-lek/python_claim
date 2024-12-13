#!/usr/bin/env python
# coding: utf-8

# In[18]:


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select


import requests
import shutil
import os
import pandas as pd
import time
from urllib.parse import quote
from time import sleep
import datetime
import pymysql 
from configparser import ConfigParser

# ติดตั้ง ChromeDriver โดยอัตโนมัติและเริ่ม WebDriver

driver = webdriver.Chrome()
configur = ConfigParser()
#configur.read('C:\\python_claim\\config.inc')
basepatch = os.path.abspath(os.getcwd())
basepatch = basepatch+'\\'+'config.inc'
configur.read(f'{basepatch}')

#driver = webdriver.Chrome()

username = str(configur.get('eclaiminfo','username'))
password = str(configur.get('eclaiminfo','passname'))

# เข้าถึงเว็บไซต์ Shopee

driver.get('https://eclaim.nhso.go.th/webComponent/main/MainWebAction.do')

# ค้นหา input field โดยใช้ชื่อ (name) และกรอกข้อมูล
input_element = driver.find_element(By.NAME, 'user')
input_element.send_keys(username)
# ค้นหา input field สำหรับรหัสผ่าน
input_element_password = driver.find_element(By.NAME, 'pass')
input_element_password.send_keys(password)


# ใช้ WebDriverWait เพื่อรอให้ element โหลดและพร้อมสำหรับการคลิก
submit_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/center/table/tbody/tr[3]/td[1]/div/div[1]/form/table/tbody/tr[3]/td/input[1]'))
)

submit_button.click()


###driver.get('https://eclaim.nhso.go.th/webComponent/validation/ValidationMainAction.do?maininscl=ucs')

##driver.get('https://eclaim.nhso.go.th/webComponent/validation/ValidationMainAction.do?maininscl=lgo')

driver.get('https://eclaim.nhso.go.th/webComponent/validation/ValidationMainAction.do?maininscl=ofc')


##เลือกเดือนที่ต้องการโหลด
###options_mo = (By.ID, 'mo')
###select_mo = Select(WebDriverWait(driver, 10).until(EC.element_to_be_clickable(options_mo)))
###select_mo.select_by_value("4")

###options_ye = (By.ID, 'ye')
###select_ye = Select(WebDriverWait(driver, 10).until(EC.element_to_be_clickable(options_ye)))
###select_ye.select_by_value("2567")
###driver.find_element(By.XPATH, "//input[@type='submit']").click()

table_element = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/center/table/tbody/tr[3]/td[3]/div[2]/table/tbody/tr[5]/td/table[1]'))
)

#ฟังชั่นหาไฟล์ซ้ำ
def check_file_exits(fname):
    conn = pymysql.connect(
        host=configur.get('servfordb','host') ,
        user=configur.get('servfordb','user'),
        password=configur.get('servfordb','pwsd'),
        db='loeifns'
    )
    cur = conn.cursor()
    select_file = "SELECT * FROM pkrep WHERE concat(documents_name,'.ecd') = %(file_name)s"
    cur.execute(select_file, { 'file_name': fname })
    return cur.rowcount


# ค้นหาแถวทั้งหมดภายใน table
rows = table_element.find_elements(By.XPATH, './/tbody/tr')
number_of_rows = len(rows)
row1= 0
counter=0
items_rows = 0
new_file = 0
# ลูปผ่านแถวทั้งหมดและดึงค่าของ <td> ที่ต้องการ
for index, row in enumerate(rows, start=1):  # ใช้ enumerate เพื่อเพิ่มลำดับ
    td_elements = row.find_elements(By.XPATH, './/td')
    if len(td_elements) > 0:  # ตรวจสอบว่ามี REP 
        rep_no = td_elements[1].text  # แสดง REP_NO
        #rep_no # แสดง REP_NO
        file_name = td_elements[4].text   #  แสดง file_name
        a_element = td_elements[13].find_element(By.XPATH, './/a') # ค้นหาลิงก์ภายใน <td> โดยใช้ XPath .//a เพื่อหาลิงก์
        file_download_url = a_element.get_attribute('href')  # ดึง URL ของลิงก์
        #file_download_url
        items_rows += 1
        if check_file_exits(file_name) ==0 :
            a_element.click() # คลิกที่ลิงก์เพื่อดาวน์โหลดไฟล์
            counter +=1
            new_file +=1 
            print(f'ลำดับที่ : {items_rows} ครั้งที่: {counter} ชื่อไฟล์  :{file_name}')
            if counter > 5:
                for i in range(4):
                    driver.execute_script("window.scroll(" + str(i * 6000) + ", " + str((i + 1) * 6000) + ")")
                    sleep(3)
                    counter=0
            
        else:
            print(f'ลำดับที่ : {items_rows} มีไฟล์แล้ว :{file_name}  ')

print(f'จำนวนแถวในตารางคือ: {number_of_rows} แถว' )
# ปิดเบราว์เซอร์เมื่อเสร็จสิ้น
# driver.quit()
sleep(5)
driver.close()


if new_file > 0 : 
    # เส้นทางของไฟล์ต้นฉบับและที่เก็บใหม่
    sourcepatch = configur.get('fileupload','source')
    destinationpatch = configur.get('fileupload','destinationOFC')
    source = r''+sourcepatch+''  ## โฟรเดอต้นทาง
    destination = r''+destinationpatch+''    ## โฟรเดอปลายทาง
    # gather all files
    allfiles = os.listdir(source)
    
    # iterate on all files to move them to destination folder
    for f in allfiles:
        src_path = os.path.join(source, f)
        dst_path = os.path.join(destination, f)
        shutil.move(src_path, dst_path)
        sleep(2)
        print(f'ย้ายไฟล์ : {f} สำเร็จ' )
    
    driver = webdriver.Chrome()
    serv = configur.get('servforwebapp','ip')
    driver.get('http://'+serv+'/repimport/auto_test.php?nhso=ofc')
    sleep(60)
    driver.close()


