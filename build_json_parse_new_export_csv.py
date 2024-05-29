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

title_line = ['日期','類型','主題','連結','內容','講者','代號','標籤','時長','下個主題1','下個主題2','下個主題3','下個主題4','下個主題5','下個主題6','下個主題7','下個主題8','下個主題_1','下個主題_2','下個主題_3','下個主題_4','下個主題_5','下個主題_6','下個主題_7','下個主題_8','相關1','相關2','相關3','相關4','相關5','相關6','相關7','相關_1','相關_2','相關_3','相關_4','相關_5','相關_6','相關_7','影片','Session type', 'Topic', 'Level', 'Delivery Type', 'Tag', 'Recording Availability', 'Programming Language', 'Session type_', 'Topic_', 'Level_', 'Delivery Type_', 'Tag_', 'Recording Availability_', 'Programming Language_','resource_0', 'Download Video', 'Download Transcript', 'BRK146', 'BRK151', 'Learn more about AKS announcemen...', 'Code to cloud with AKS and GitHub', 'Assess your cloud native maturit...', 'GitHub Copilot', 'Visual Studio 17.6', 'Microsoft Dev Box', 'AI Toolkit', '.NET Aspire', 'Semantic Kernel and Astra DB', 'RAG Explained', 'Astra DB documentation', 'GitHub Copilot - Your AI pair pr...', 'What is GitHub Copilot?', 'GitHub Copilot video series', 'GitHub Certifications', 'Collection: Accelerate Developer...', 'BRK143', 'BRK210', 'DEM728', 'LAB320', 'BRK203', 'Accelerate Development', 'Dev Box Resources', 'Deploy an application that uses ...', 'BRK150', 'GitHub Advanced Security Overvie...', 'GitHub Advanced Security Docs', 'Collection: Secure Developer Pla...', 'Snowflake & Azure OpenAI', 'More on this topic', 'Build custom GPT apps for all pl...', '.NET AI Collection', 'New .NET AI Learn Modules', '.NET Beginner Videos', 'Codespaces Quickstart', 'Codespaces deep-dive', 'AI Search now offers more capaci...', 'BRK148', 'Get started with Azure Kubernete...', 'Learn more about ClickHouse', 'Build AI Apps with Azure Databas...', 'Azure Database for PostgreSQL - ...', 'Collection: Microsoft Developer ...', 'GitHub Features - check out all ...', 'LAB341 Collection', '.NET Aspire blog', '.NET Aspire Webpage', '.NET Aspire Collection', 'View the guidance link', 'Join the Arm Developer Program', 'Find out more about Arm Learning...', 'Announcing new search capabiliti...', 'Learn .Net Aspire', '.NET Aspire Blog', "What's New with Assistant's", 'Learn .NET Aspire', '.NET Aspire Videos', 'Get Started with C# and .NET', 'Build an end to end RAG system w...', 'Learn more about LLMOps in Azure AI', 'NGINXaaS: Azure Native Service', 'Get started: NGINXaaS', 'Accessibility config sample', 'Customizations sample', 'Customizations blog post', 'Build announcements', 'Windows Dev Center', 'WinUI 3 Documentation', 'WPF Documentation', 'WinUI Community Call', 'So, what is GitHub?', 'BRK141', 'LAB301', 'Dev Box resources', 'Learn about new C# features', 'Track feature implementation status', 'Participate in C# language design', 'BRK152', 'Blazor Hybrid Docs', 'Download Visual Studio 2022 17.1...', 'Build a .NET MAUI Blazor Hybrid ...', 'Overview of Microsoft Defender f...', 'Learn more about platform engine...', 'Explore careers at H&M', 'Azure Deployment Environments do...', 'New extensibility model in Azure...', 'Azure Deployment Environments', 'What is Azure DevTest Labs?', 'Azure Deployment Environments An...', '.New .NET AI Learn Modules', 'Run Oracle Databases at Azure', 'Vidoes to learn how it works?', 'Overview of Microsoft Informatio...', 'Flink datasheet', 'Flink 101', 'Blog', 'Tech Docs', 'Overview Microsoft Purview Infor...', 'Azure Load Testing', 'Playywright Testing', 'Playwright Testing Preview Docum...', 'Apps on Azure Blog', 'Python Risk Identification Tool ...', 'Microsoft’s open automation fram...', 'Microsoft AI Red Team building f...', 'AI code generation: benefits', 'WinUI 3 Gallery App', 'WPF Gallery App', 'Learn more about Microsoft for S...', 'Creating GenAI Experiences with ...', 'ISV Success through the ISV Hub', 'Blogs', 'Migrate Workloads: F5 NGINXaaS', 'Get started with NGINXaaS', 'Session GitHub', 'ASP.NET Core Learn Docs', 'Web API Development in VS', 'Safely use secrets in HTTP Requests', 'Revolutionize your development', 'Session Slide Deck', 'GSMA Showcase Page', 'MS Dynamics Test Automation', '.NET diagnostic tools', '.NET Monitor docs and samples', 'StackHawk Overview & Demo', 'Videos to learn how it works', 'Document Viewer for .NET', 'Learn .NET MAUI', '.NET MAUI Repo', 'Azure Cosmos DB Conf 2024 Session', '.NET Conf 2023 Session', 'What’s New in EF9', 'Community Standup', 'MongoDB provider', 'Elasticsearch and Azure OpenAI', 'Elasticsearch & OpenAI Configs', 'View slide deck', 'Book a demo', '.NET Packages on Ubuntu 24.04', 'Get Started', 'Datasheet', 'Case Studies', 'MadeWithWisej.com', 'Free Analysis', 'Blazor Home', '.NET 9', '.NET Smart Components', 'Data Sheet NeuVector', 'Data Sheet Rancher Prime', 'Data Streaming for AI', 'Confluent + Azure Blog Post', 'Learn more about ReSharper', 'Try JetBrains AI Assistant', 'ADF and Snowflake', 'Quickstart: Power Apps', 'Connector: Power Apps', 'Quickstart: Azure Data Factory', 'WinForms GitHub', "Integrating Isovalent's Cilium", 'Isovalent in AKS', 'AI Code generation: Pros& Cons', '附註']
block_pills = []
tag_fields = []
session_next_step_count = 0
related_session_count = 0
speaker_counts = 0

new_build_lists = []
with open(os.path.join(src_path, 'build_csv_with_download.csv'), 'w', encoding='utf-8-sig', newline = '') as csv_f:
    csv_writer = csv.writer(csv_f)
    csv_writer.writerow(title_line)
    for i, build_list in enumerate(build_lists):
        if build_list[0] == '日期':
            new_build_lists.append(title_line)
            continue

        print(f'{i}.', end='')

        session_page = build_list[title_line.index('連結')]
        title_file_name = build_list[title_line.index('主題')]
        title_file_name = re.sub(r'[^\w\s-]', '', title_file_name.lower())
        title_file_name = re.sub(r'[-\s]+', '-', title_file_name).strip('-_')
        title_file_name = f'{i}_{title_file_name}.json'

        with open(os.path.join(src_path, data_path, title_file_name), 'r', encoding='utf-8') as f_title:
            element = f_title.read()

        new_build_list = [''] * len(title_line)
        new_build_list[title_line.index('日期')] = build_list[title_line.index('日期')]
        new_build_list[title_line.index('類型')] = build_list[title_line.index('類型')]
        new_build_list[title_line.index('主題')] = build_list[title_line.index('主題')]
        new_build_list[title_line.index('連結')] = build_list[title_line.index('連結')]
        new_build_list[title_line.index('內容')] = build_list[title_line.index('內容')]

        if not element.startswith('{'):
            new_build_list[title_line.index('附註')] = element
        else:

            element = json.loads(element)

            session_code = ''
            if element.get('session_code'):
                session_code = element.get('session_code')
            new_build_list[title_line.index('代號')] = session_code

            speakers = ''
            if element.get('speakers'):
                speakers = element.get('speakers')
                if len(speakers) > 0:
                    speakers = ','.join(speaker for speaker in speakers)
            new_build_list[title_line.index('講者')] = speakers

            block_pill = ''
            if element.get('block_pills'):
                blocks = element.get('block_pills')
                if len(blocks)!= 1:
                    print('error! blocks')
                    break
                block_pill = ''.join(pill for pill in blocks[0])
            new_build_list[title_line.index('標籤')] = block_pill

            duration_item = ''
            if element.get('duration_items'):
                duration_items = element.get('duration_items')
                if len(duration_items) != 1:
                    print('error! duration_items')
                    break
                duration_item = duration_items[0]
            new_build_list[title_line.index('時長')] = duration_item.replace('Duration ', '')

            if element.get('session_next_steps'):
                next_steps = element.get('session_next_steps')
                if len(next_steps) > 0:
                    for j, step in enumerate(next_steps):
                        for key, value in step.items():
                            new_build_list[title_line.index(f'下個主題{j+1}')] = f'{key}'
                            new_build_list[title_line.index(f'下個主題_{j+1}')] = f'{value}'

            if element.get('related_sessions'):
                related_sessions = element.get('related_sessions')
                if len(related_sessions) > 0:
                    for k, step in enumerate(related_sessions):
                        for key, value in step.items():
                            new_build_list[title_line.index(f'相關{k+1}')] = f'{key}'
                            new_build_list[title_line.index(f'相關_{k+1}')] = f'{value}'

            if element.get('video_detail'):
                new_build_list[title_line.index(f'影片')] = element.get('video_detail')

            if element.get('tag_fields'):
                tags = element.get('tag_fields')
                if len(tags) > 0:
                    for key0, value0 in tags.items():
                        for key1, value1 in value0.items():
                            new_build_list[title_line.index(f'{key0}')] = f'{key1}'
                            new_build_list[title_line.index(f'{key0}_')] = f'{value1}'

            if element.get('resources'):
                resources = element.get('resources')
                if len(resources) > 0:
                    for key0, value0 in resources.items():
                        new_build_list[title_line.index(key0)] = f'{value0}'

            if element.get('speaker_blocks'):
                speaker_block = ''
                speaker_blocks = element.get('speaker_blocks')
                for block in speaker_blocks:
                    speaker_block += block
                new_build_list[title_line.index('附註')] = f'{speaker_block}'

        new_build_lists.append(new_build_list)
        csv_writer.writerow(new_build_list)

print(f'== end ==')



