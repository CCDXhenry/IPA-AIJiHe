import pandas as pd
import logging


######主程序######
#获取数据表
logging.info("开始读取数据表...调剂数据文件获取")
df_生成文件信息表 = pd.read_csv(file_path_生成文件信息表)
logging.info("数据表读取完成")
#获取文件信息
all_file_info = []
for index, row in df_生成文件信息表.iterrows():
    薪酬类型 = row['薪酬类型']
    统计月份 = row['统计月份']
    工单级别 = row['工单级别']
    政策分期小项编码 = row['政策分期小项编码']
    渠道归属县市 = row['渠道归属县市']
    文件路径 = row['文件路径']
    file_info = {
        '薪酬类型': 薪酬类型,
        '统计月份': 统计月份,
        '工单级别': 工单级别,
        '政策分期小项编码': 政策分期小项编码,
        '渠道归属县市': 渠道归属县市,
        '文件路径': 文件路径
    }
    all_file_info.append(file_info)