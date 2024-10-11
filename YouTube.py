import os
import time

import pyautogui
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.common.by import By

import yt_dlp

import pandas as pd
import openpyxl
from openpyxl import load_workbook

from local_log import logger

class YouTube:
    def __init__(self):

        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
        chrome_options.add_argument("--disable-blink-features")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 6.1; WOW64; rv:30.0) Gecko/20100101 Firefox/30.0')
        s = Service('./chromedriver.exe')
        self.browser = webdriver.Chrome(service=s, options=chrome_options)

        # TODO: Upload your own cookies by referring to README.md
        cookie_file = './cookies.txt'
        ydl_opts = {
            'cookiefile': cookie_file,
            'quiet': True,
            'extract_flat': True,
        }

    def set_keyword(self, keywords:str):
        self.keyword = keywords

    def search_video(self, uploader, channel_name, file_path, scrollTime:int):
        # searchUrl = "https://www.youtube.com/results?search_query={}&sp=EgIQAQ%253D%253D".format(self.keyword)
        searchUrl = "https://www.youtube.com/{}/search?query={}".format(channel_name, self.keyword)
        self.browser.get(searchUrl)
        self.browser.maximize_window()
        time.sleep(5)

        self.__scroll2bottom(scrollTime)

        video_cards = self.browser.find_elements(By.TAG_NAME, 'ytd-video-renderer')

        videos_list = []
        for card in video_cards:
            title = card.find_element(By.ID, 'video-title').get_attribute("title")
            url = card.find_element(By.ID, 'video-title').get_attribute("href")

            # try:
            #     video = self.get_video_information(title, url)
            # except Exception as e:
            #     logger.error(title + str(e))

            '''
                TODO: Since there may be problems when crawling video information, 
                it is recommended to store the url of the video first, and then 
                crawl the corresponding video information
            '''
            video = {"uploader": uploader, "title": title, "url_article": url}

            videos_list.append(video)

        df_new = pd.DataFrame(videos_list)

        if os.path.exists(file_path):
            with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                startrow = writer.sheets['Sheet1'].max_row
                df_new.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=startrow)
        else:
            df_new.to_excel(file_path, index=False, sheet_name='Sheet1')

    def get_video_information(self, title, url):
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=False)

            video = {"uploader": info.get('uploader'),
                    "title": title,
                    "link": url,
                    "published_date": info.get('upload_date'),
                    "duration": info.get('duration'),
                    "views": info.get('view_count'),
                    "likes": info.get('like_count'),
                    "comments": info.get('comment_count')}
        return video

    def __scroll2bottom(self, scrollTime:int):
        for i in range(scrollTime):
            pos = pyautogui.size()
            pyautogui.moveTo(pos.width / 2, pos.height / 2)
            pyautogui.scroll(-400)
            time.sleep(0.2)
