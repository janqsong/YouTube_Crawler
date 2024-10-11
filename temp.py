import os
import yt_dlp

import pandas as pd
import time

from local_log import logger

cookie_file = './www.youtube.com_cookies (1).txt'

# 读取Excel文件
df = pd.read_excel(f'./youtube_other_process.xlsx', engine='openpyxl')

items_list = df.values.tolist()

ydl_opts = {
    'cookiefile': cookie_file,
    'quiet': True,
    'extract_flat': True,  # 获取视频信息而不是下载视频
}

videos_list = []
try:
    for i in range(2449, len(items_list)):
        item = items_list[i]
        title = item[0]
        url = item[1]

        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=False)
            channel = info.get('uploader')
            publish_date = info.get('upload_date')
            duration = info.get('duration')
            view_count = info.get('view_count')
            like_count = info.get('like_count')
            comment_count = info.get('comment_count')
        if comment_count is None:
            comment_count = 0

        video = {"channel": channel,
                "title": title,
                "link": url,
                "published_date": publish_date,
                "duration": duration,
                "views": view_count,
                "likes": like_count,
                "comments": comment_count}
        logger.info(str(i) + str(video))
        print(i, video)
        videos_list.append(video)
        time.sleep(1)
except Exception as e:
    logger.info(i)
    logger.info(e)
    print(i, title)

file_path = './youtube_final.xlsx'
df_new = pd.DataFrame(videos_list)

if os.path.exists(file_path):
    # 如果文件存在，追加到 Sheet1
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        # 计算已有数据的行数
        startrow = writer.sheets['Sheet1'].max_row
        df_new.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=startrow)
else:
    # 如果文件不存在，创建一个新的文件
    df_new.to_excel(file_path, index=False, sheet_name='Sheet1')

