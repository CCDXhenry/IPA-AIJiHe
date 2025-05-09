import os
from datetime import datetime, timedelta
import logging

# 获取当前年月
current_date = datetime.now()
year_suffix = str(current_date.year)[2:]  # 获取年份后两位
month = current_date.month
day = current_date.day
project_path = rf"{basic_path}\{year_suffix}年{month}月{day}日数据"
if not os.path.exists(project_path):
    os.makedirs(project_path)
    logging.info("文件夹创建成功")


file_path_费用结算数据稽核 = rf"{project_path}\关于{year_suffix}年{month}月费用结算数据稽核.xlsx"
file_path_调剂数据导入模板 = rf"{project_path}\调剂数据导入模板.xls"
file_path_调剂依据文件 = rf"{project_path}\调剂依据文件.xlsx"

file_path_调剂表数据目录 = rf"{project_path}\调剂表数据目录"
if not os.path.exists(file_path_调剂表数据目录):
    os.makedirs(file_path_调剂表数据目录)
    logging.info("文件夹创建成功")
    
file_path_生成文件信息表 = rf"{file_path_调剂表数据目录}\生成文件信息表.csv"

file_path_调剂数据明细取数目录 = rf"{project_path}\调剂数据明细取数目录"
if not os.path.exists(file_path_调剂数据明细取数目录):
    os.makedirs(file_path_调剂数据明细取数目录)
    logging.info("文件夹创建成功")

file_path_调剂数据明细取数表 = rf"{file_path_调剂数据明细取数目录}\调剂数据明细取数表.xlsx"
file_path_调剂明细数据 = rf"{project_path}\调剂明细数据.xlsx"

file_path_调剂金额核验结果 = os.path.join(project_path, '调剂金额核验结果.xlsx')

logging.info("文件初始化完成")