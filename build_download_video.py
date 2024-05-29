# author: jacques.spectre@gmail.com
# date: 2022-04-26, 12:23pm
# Jacques 工作室
# all rights are reserved
# version: 0.1
# from parsel import Selector
# import requests

from datetime import datetime
import random
import requests
import time
import json
import csv
import sys
import os
import re

# 指定 CSV 檔案路徑
csv_file_path = 'build_video_csv.csv'
downloads = '.downloads'
src_path = os.getcwd()

headers = {
    'Accept': '*/*',
    'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
    'Connection': 'keep-alive',
    # 'Cookie': 'MC1=GUID=d6ed8acba4494aba9aa7afc04f76b8cf&HASH=d6ed&LV=202405&V=4&LU=1715047777425; MSCC=NR; ak_bmsc=2BF16F32CB354C9D7433834EA1FE5C67~000000000000000000000000000000~YAAQv9fSF4tfgKSPAQAAqLFCtRdPazugjNgOJJVmmUdJU1s4eYLIBV4Dfz7kk0bwTmMxFUVFNvC5tJQYGDWNtCvXWMAjYN3VEMCW11K++ciDhdz9aA5GcRTtr8E1raYJPoZET4WDCGINXTpJ0Tvxi5jqAJrdvkPEkBM7+UZ2ao6FTOPyU2Swtbv+2+/vCHw/fFFAi3BeH1CdFzBIZ6KC0Kl83odKDFYVlrI8iYRopti2Vq0uDCJjQra9PL9IQ2hC4bKzD6rLP0L2H9vXi0ihQ9nY0i5LkBLt+gjOk9npkc5SvCFwtm6YMbMw/Qg62veuPw6NX0SvihDN5QtTUV+HikEFirGfCQ0a/tuSPS0Q1MEcfMdPPt2U/6XMWCY3IYADY0EzScIOY+5QTpv2HsiTAFBn8qylwbmJNfhiR3msik/Txg==',
    'DNT': '1',
    'If-Range': '"0x8DC79FECDFB7293"',
    'Range': 'bytes=206536704-1268159212',
    'Referer': 'https://mediusdownload.event.microsoft.com/asset-2e41c715-00d5-4f79-ba9a-161a4438aab5/Video-BRK165_v1-FHD-6000kbps-6000.mp4?sv=2018-03-28&sr=b&sig=uZWgVWPAIlJyvr4u2eBmFP6qZVvjjQtMl1aYsFl0vbU%3D&st=2024-05-22T01%3A26%3A04Z&se=2029-05-22T01%3A31%3A04Z&sp=r&rscd=filename%3DN2F9-BRK165-Microsoft%2BFabric%253a%2BWhat%2527s%2Bnew%2Band%2Bwhat%2527s%2Bnext.mp4',
    'Sec-Fetch-Dest': 'video',
    'Sec-Fetch-Mode': 'no-cors',
    'Sec-Fetch-Site': 'same-origin',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36 Edg/125.0.0.0',
    'X-Edge-Shopping-Flag': '1',
    'sec-ch-ua': '"Microsoft Edge";v="125", "Chromium";v="125", "Not.A/Brand";v="24"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
}

# 定義下載函數
def download_file(url, filename):
    try:
        with requests.get(url, headers=headers) as response:
            response.raise_for_status()  # 確保請求成功
            with open(filename, 'wb') as f:
                for chunk in response.iter_bytes():
                    f.write(chunk)
        print(f'Successfully downloaded {filename}')
    except Exception as e:
        print(f'Failed to download {filename}: {e}')


# 讀取 CSV 並下載檔案
with open(os.path.join(src_path, csv_file_path), 'r', newline='', encoding='utf-8-sig') as csvfile:
    reader = csv.DictReader(csvfile)
    for i, row in enumerate(reader):
        video_url = row.get('Download Video', '')
        transcript_url = row.get('Download Transcript', '')
        title_file_name = row.get('Subject', '')
        title_file_name = re.sub(r'[^\w\s-]', '', title_file_name.lower())
        title_file_name = re.sub(r'[-\s]+', '-', title_file_name).strip('-_')
        title_file_name = f'{i+1}_{title_file_name}'


        if video_url:
            video_filename = os.path.join(src_path, '.downloads', f'{title_file_name}.mp4')
            print(f'=== {i + 1}. {video_filename}', end='')
            download_file(video_url, video_filename)

        if transcript_url:
            transcript_filename = os.path.join(src_path, '.downloads', f'{title_file_name}.docx')
            print(f'=== {i + 1}. {transcript_filename}', end='')
            download_file(transcript_url, transcript_filename)






print(f'== end ==')



