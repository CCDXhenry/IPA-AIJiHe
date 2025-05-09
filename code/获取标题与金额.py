import os
import pandas as pd
import logging
from datetime import datetime
#数据表
basic_path = rf"C:\project\python_project\AI+市场工具&稽核\AI+市场工具&稽核第一期需求\选项目-制表-得出结论"
# 获取当前年月
current_date = datetime.now()
year_suffix = str(current_date.year)[2:]  # 获取年份后两位
month = current_date.month
day = current_date.day
project_path = rf"{basic_path}\{year_suffix}年{month}月{day}日数据"
if not os.path.exists(project_path):
    os.makedirs(project_path)
    logging.info("文件夹创建成功")

# 设置日志
logging.info("开始处理数据...")
file_path_费用结算数据稽核 = rf"{project_path}\关于{year_suffix}年{month}月费用结算数据稽核.xlsx"

# 获取file_path_费用结算数据稽核文件名称
work_title = os.path.basename(file_path_费用结算数据稽核)
work_title = work_title.split('.')[0]

# 获取汇总结算金额
df_汇总结算金额 = pd.read_excel(file_path_费用结算数据稽核, sheet_name='汇总结算金额')
# 获取总金额
work_amount = df_汇总结算金额.iloc[-1]['结算金额']
