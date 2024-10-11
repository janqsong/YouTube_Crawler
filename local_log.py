import logging

logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

file_handler = logging.FileHandler('output.log', mode='a', encoding='utf-8')
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)