import pandas as pd
import logging
from datetime import datetime
import os

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

file_path_待调剂明细数据 = rf"{project_path}\待调剂明细数据.xlsx"
file_path_制表取数表 = rf"{project_path}\制表取数表.xlsx"

#读取file_path_制表取数表【文件路径】，将所有包含的文件整合生成file_path_待调剂明细数据
df_制表取数表 = pd.read_excel(file_path_制表取数表)
df_待调剂明细数据 = pd.DataFrame()
try:
    for i, row in df_制表取数表.iterrows():
        file_path = row["文件路径"]
        if pd.notnull(file_path):
            df = pd.read_excel(file_path)
            df_待调剂明细数据 = pd.concat([df_待调剂明细数据, df], ignore_index=True)
except Exception as e:
    logging.error(f"读取文件时发生错误：{e}")
try:
    df_待调剂明细数据.to_excel(file_path_待调剂明细数据, index=False)
    logging.info("待调剂明细数据整合完成")
except Exception as e:
    logging.error(f"保存待调剂明细数据时发生错误：{e}")