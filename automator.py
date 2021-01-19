#!/usr/bin/python
# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, \
    UnexpectedAlertPresentException, NoAlertPresentException
from time import sleep
from urllib.parse import quote
from sys import platform
import pandas

options = Options()
driver = webdriver.Chrome('drivers\\chromedriver.exe', options=options)
driver.get('https://web.whatsapp.com')
count = 0
delay = 30

excel_data = pandas.read_excel('JRR Marathon - King ِAbdulAziz Road (Responses).xlsx', sheet_name='Responses')
input('Press ENTER after login into Whatsapp Web and your chats are visiable.'
      )
for column in excel_data['Full Name'].tolist():

    # takes message from excel

    message = excel_data['Arb Message'][0]
    message = message.replace('{بيب}',
                              str(excel_data['BIB Number'][count]))
    message = message.replace('{arabic_run_distance}',
                              str(excel_data['Arabic Run Distance'
                              ][count]))

    # message = message.replace('{bib_ number}', str(excel_data['BIB Number'][count]))
    # message = message.replace('{run_distance}', str(excel_data['Run Distance'][count]))

    # take numbers from excel

    number = str(excel_data['Phone Number'][count])
    count = count + 1

    try:
        url = 'https://web.whatsapp.com/send?phone=' + number \
            + '&text=' + message
        driver.get(url)
        try:
            click_btn = WebDriverWait(driver,
                    delay).until(EC.element_to_be_clickable((By.CLASS_NAME,
                                 '_2Ujuu')))
        except (UnexpectedAlertPresentException,
                NoAlertPresentException), e:
            print 'alert present'
            Alert(driver).accept()
        sleep(1)
        click_btn.click()
        sleep(3)
        print 'Message sent to: ' + number
    except Exception, e:
        print 'Failed to send message to ' + number + str(e)
