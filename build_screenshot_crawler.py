# author: jacques.spectre@gmail.com
# date: 2022-04-26, 12:23pm
# Jacques 工作室
# all rights are reserved
# version: 0.1
# from parsel import Selector
# import requests

from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, ElementClickInterceptedException, UnexpectedAlertPresentException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium import webdriver
from dateutil.relativedelta import relativedelta
from bs4 import BeautifulSoup
from datetime import datetime
import random
import time
import json
import csv
import sys
import os
import re

build_csv = 'Microsoft Build.csv'
src_path = os.getcwd()
maxWaitInSec = 10  # 秒
randomSecs = [2, 3]
data_path = './.screenshots_collection'
if not os.path.isdir(os.path.join(src_path, data_path)):
    os.makedirs(os.path.join(src_path, data_path))


with open(os.path.join(src_path, build_csv), 'r', encoding='utf-8-sig') as f:
    build_lists = list(csv.reader(f))

title_line = None
counter = 0

option = webdriver.FirefoxOptions()
option.add_argument('--headless')
option.add_argument('--disable-gpu')
option.add_argument("--auto-open-devtools-for-tabs")

browser = webdriver.Firefox(options=option)
print()

for i, build_list in enumerate(build_lists):
    if build_list[0] == '日期':
        title_line = build_list
        continue

    if not title_line:
        break

    #if i<11:
    #    continue

    session_page = build_list[title_line.index('連結')]
    title_file_name = build_list[title_line.index('主題')]
    title_file_name = re.sub(r'[^\w\s-]', '', title_file_name.lower())
    title_file_name = re.sub(r'[-\s]+', '-', title_file_name).strip('-_')
    title_file_name = f'{i}_{title_file_name}.png'

    try:

        browser.get(session_page)

        time.sleep(random.choice(randomSecs))

        WebDriverWait(browser, timeout=maxWaitInSec).until(EC.presence_of_element_located(
            (By.CLASS_NAME, 'session-detail-ignite__sidebar-section')) and EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, '#video')) and EC.presence_of_all_elements_located((By.CLASS_NAME,'session-block__session-code')))

        scroll_width = browser.execute_script('return document.body.parentNode.scrollWidth')
        scroll_height = browser.execute_script('return document.body.parentNode.scrollHeight')
        browser.set_window_size(scroll_width, scroll_height)


    except Exception as e:
        print('頁面開啟錯誤！')
        element = str(e)

    else:
        try:
            browser.save_full_page_screenshot(os.path.join(src_path, data_path, title_file_name))
        except Exception as e:
            print('頁面擷取錯誤！')
            element = str(e)

    print(f'== 擷取 {i}==> {title_file_name} 結束! ==')
    time.sleep(random.choice(randomSecs))
browser.quit()



print(f'== 程式結束! ==')