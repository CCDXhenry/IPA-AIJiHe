import pandas as pd
import logging
from datetime import datetime
import os

#读取file_path_制表取数表【文件路径】，将所有包含的文件整合生成file_path_调剂明细数据
df_调剂数据明细取数表 = pd.read_excel(file_path_调剂数据明细取数表)
df_调剂明细数据 = pd.DataFrame()
try:
    for i, row in df_调剂数据明细取数表.iterrows():
        file_path = row["文件路径"]
        if pd.notnull(file_path):
            df = pd.read_excel(file_path)
            df_调剂明细数据 = pd.concat([df_调剂明细数据, df], ignore_index=True)
except Exception as e:
    logging.error(f"读取文件时发生错误：{e}")
try:
    df_调剂明细数据.to_excel(file_path_调剂明细数据, index=False)
    logging.info("待调剂明细数据整合完成")
except Exception as e:
    logging.error(f"保存待调剂明细数据时发生错误：{e}")