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
data_path = './.data_collection'
if not os.path.isdir(os.path.join(src_path, data_path)):
    os.makedirs(os.path.join(src_path, data_path))


with open(os.path.join(src_path, build_csv), 'r', encoding='utf-8-sig') as f:
    build_lists = list(csv.reader(f))

title_line = None
counter = 0
browser = webdriver.Firefox()

for i, build_list in enumerate(build_lists):
    if build_list[0] == '日期':
        title_line = build_list
        continue

    if not title_line:
        break

    if i<11:
        continue

    session_page = build_list[title_line.index('連結')]
    title_file_name = build_list[title_line.index('主題')]
    title_file_name = re.sub(r'[^\w\s-]', '', title_file_name.lower())
    title_file_name = re.sub(r'[-\s]+', '-', title_file_name).strip('-_')
    title_file_name = f'{i}_{title_file_name}.json'

    try:

        browser.get(session_page)

        time.sleep(random.choice(randomSecs))

        WebDriverWait(browser, timeout=maxWaitInSec).until(EC.presence_of_element_located(
            (By.CLASS_NAME, 'session-detail-ignite__sidebar-section')) and EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, '#video')) and EC.presence_of_all_elements_located((By.CLASS_NAME,'session-block__session-code')))

    except Exception as e:
        print('頁面開啟錯誤！')
        element = str(e)

    else:
        try:
            resources = browser.find_elements(By.CLASS_NAME, 'resource-button__link')
            resource_items = {}
            resource_count = 0
            for resource in resources:
                resource_key = resource.text
                if not (resource_key):
                    resource_key = f'resource_{resource_count}'
                    resource_count += 1
                resource_items[resource_key] = resource.get_attribute('href')
            tag_fields = browser.find_elements(By.CLASS_NAME, 'tags__field')
            tag_items = {}
            tag_count = 0
            for tag in tag_fields:
                tag_title = tag.find_element(By.CLASS_NAME, 'tags__field-title').text
                tag_content = tag.find_element(By.CLASS_NAME, 'dv-blink__content').text
                tag_href = tag.find_element(By.TAG_NAME, 'a').get_attribute('href')
                tag_items[tag_title] = {
                    tag_content: tag_href
                }

            next_steps = []
            next_step_buttons = browser.find_elements(By.CLASS_NAME, 'resource-button')
            for next_button in next_step_buttons:
                button_content = {}
                button_href = next_button.find_element(By.TAG_NAME, 'a').get_attribute('href')
                button_labels = next_button.parent.find_elements(By.ID, 'nextStepLabel')
                if len(button_labels)>0:
                    if len(button_labels)!=1:
                        print(f'len button_labels == {len(button_labels)}')
                        continue
                    button_label = button_labels[0].text
                    button_content[button_label] = button_href
                    next_steps.append(button_content)

            session_code = browser.find_element(By.CLASS_NAME,'session-block__session-code').text
            element = {
                'session_code' : session_code,
                'block_pills' : [pill.text.split('\n') for pill in browser.find_elements(By.CLASS_NAME,'session-block__pills')],
                'speakers' : [speaker.text.replace('\n', '') for speaker in browser.find_elements(By.CLASS_NAME, 'person-badge')],
                'duration_items' : [duration.text for duration in browser.find_elements(By.CLASS_NAME,'session-details-header-banner__duration')],
                'session_next_steps' : next_steps,
                'related_sessions' : [{session.text: session.get_attribute('href')} for session in browser.find_elements(By.CLASS_NAME, 'session-block__link')],
                'video_detail' : browser.find_element(By.CSS_SELECTOR, '#video').text,
                'tag_fields' : tag_items,
                'resources' : resource_items,
                'speaker_blocks' : [speaker.text for speaker in browser.find_elements(By.CLASS_NAME, 'speaker-block')],
                'orig_content':  build_list,
            }

        except Exception as e:
            print('頁面擷取錯誤！')
            element = str(e)

    with open(os.path.join(src_path, data_path, title_file_name), 'w', encoding='utf-8') as f_title:
        json.dump(element, f_title, ensure_ascii=False, indent=4)

    print(f'\n== 擷取 {i}==> {title_file_name} 結束! ==')
    time.sleep(random.choice(randomSecs))
browser.quit()



print(f'== 程式結束! ==')