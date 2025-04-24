import os
from datetime import datetime, timedelta
import logging

#项目路径
date_day = datetime.today().strftime('%Y-%m-%d')
basic_path = r"C:\project\python_project\关于新入网号码回访开发需求"
# 日志配置
log_path = os.path.join(basic_path, 'logs')
if not os.path.exists(log_path):
    os.makedirs(log_path)

log_formatter = logging.Formatter('%(asctime)s %(levelname)s: %(message)s')
root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)

log_file_handler = logging.FileHandler(
    filename=os.path.join(log_path, f"{datetime.now().strftime('%Y-%m-%d')}.log"), mode='a', encoding='utf-8')
log_file_handler.setFormatter(log_formatter)
root_logger.addHandler(log_file_handler)