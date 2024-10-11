import os
import time
import pandas as pd

import yt_dlp

from YouTube import YouTube
from local_log import logger

search_keys = ['Hangzhou']
channel_lists = [['Hangzhou Feel', '@HangzhouFeel']]

def search_video(file_path):
    for channel in channel_lists:
        for search_key in search_keys:

            uploader = channel[0]
            channel_name = channel[1]

            try:
                you = YouTube()
                you.set_keyword(search_key)
                # TODO: If you need to search for more, you can turn up scrollTime.
                you.search_video(uploader, channel_name, file_path, scrollTime=0)
            except Exception as e:
                logger.error(uploader+str(e))

def get_full_information(file_path):

    # TODO: Upload your own cookies by referring to README.md
    cookie_file = './cookies.txt'
    ydl_opts = {
        'cookiefile': cookie_file,
        'quiet': True,
        'extract_flat': True,
    }

    df = pd.read_excel(file_path, engine='openpyxl')
    items_list = df.values.tolist()

    videos_list = []
    try:
        for i in range(len(items_list)):
            item = items_list[i]
            uploader = item[0]
            title = item[1]
            url = item[2]

            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(url, download=False)
                video = {"uploader": uploader,
                    "title": title,
                    "link": url,
                    "published_date": info.get('upload_date'),
                    "duration": info.get('duration'),
                    "views": info.get('view_count'),
                    "likes": info.get('like_count'),
                    "comments": info.get('comment_count')}

            logger.info(str(i) + str(video))
            print(i, video)
            videos_list.append(video)
            time.sleep(1)
    except Exception as e:
        logger.error(str(i) + str(e))
        print(i, title)

    new_file_path = "information_" + file_path
    df_new = pd.DataFrame(videos_list)

    if os.path.exists(new_file_path):
        with pd.ExcelWriter(new_file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            startrow = writer.sheets['Sheet1'].max_row
            df_new.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=startrow)
    else:
        df_new.to_excel(new_file_path, index=False, sheet_name='Sheet1')


if __name__ == "__main__":
    
    file_path = "yourname.xlsx"

    # search_video(file_path)
    get_full_information(file_path)
    