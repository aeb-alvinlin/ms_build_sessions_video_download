# author: jacques.spectre@gmail.com
# date: 2022-04-26, 12:23pm
# Jacques 工作室
# all rights are reserved
# version: 0.1
# from parsel import Selector
# import requests

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
data_path = './.data_collection'

with open(os.path.join(src_path, build_csv), 'r', encoding='utf-8-sig') as f:
    build_lists = list(csv.reader(f))

title_line = ['日期','類型','主題','連結','內容','講者','標籤','時長','下個主題1','下個主題2','下個主題3','下個主題4','下個主題5','下個主題6','下個主題7','下個主題8','相關1','相關2','相關3','相關4','相關5','相關6','相關7','影片','橌位','Session type', 'Topic', 'Level', 'Delivery Type', 'Tag', 'Recording Availability', 'Programming Language','附註']
block_pills = []
tag_fields = []
resources = []
session_next_step_count = 0
related_session_count = 0
speaker_counts = 0
new_build_list = []

for i, build_list in enumerate(build_lists):
    if build_list[0] == '日期':
        new_build_list.append(title_line)
        continue

    print(f'{i}.', end='')

    session_page = build_list[title_line.index('連結')]
    title_file_name = build_list[title_line.index('主題')]
    title_file_name = re.sub(r'[^\w\s-]', '', title_file_name.lower())
    title_file_name = re.sub(r'[-\s]+', '-', title_file_name).strip('-_')
    title_file_name = f'{i}_{title_file_name}.json'

    with open(os.path.join(src_path, data_path, title_file_name), 'r', encoding='utf-8') as f_title:
        element = f_title.read()

    element = json.loads(element)
    if isinstance(element, str):
        print(f'{i}=empty')
        continue

    if element.get('block_pills'):
        blocks = element.get('block_pills')
        for pills in blocks:
            for pill in pills:
                if not pill in block_pills:
                    block_pills.append(pill)

    if element.get('speakers'):
        speakers = element.get('speakers')
        if len(speakers) > speaker_counts:
            speaker_counts = len(speakers)

    if element.get('related_sessions'):
        related_sessions = element.get('related_sessions')
        if len(related_sessions) > related_session_count:
            related_session_count = len(related_sessions)

    if element.get('tag_fields'):
        tags = element.get('tag_fields')
        for tag in tags:
            if not tag in tag_fields:
                tag_fields.append(tag)

    if element.get('session_next_steps'):
        next_steps = element.get('session_next_steps')
        if len(next_steps) > session_next_step_count:
            session_next_step_count = len(next_steps)

    if element.get('resources'):
        resource_s = element.get('resources')
        for resource in resource_s:
            if not resource in resources:
                resources.append(resource)



print(f'\n== block_pills ==')
print(block_pills)
print(f'== tag_fields ==')
print(tag_fields)
print(f'== session_next_step_count ==')
print(session_next_step_count)
print(f'== related_session_count ==')
print(related_session_count)
print(f'== speaker_counts ==')
print(speaker_counts)
print(f'== resources ==')
print(resources)

print(f'== end ==')



