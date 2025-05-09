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

df_费用结算数据 = pd.read_excel(file_path_费用结算数据稽核, sheet_name='合作费用业务明细查询0')
#更新结算金额与调剂金额，如果结算状态为‘结算’，结算金额为‘金额’列，调剂金额为0；反之，结算金额为0，调剂金额为负的‘金额’列
df_费用结算数据['结算金额'] = df_费用结算数据.apply(lambda row: row['金额'] if row['结算状态'] == '结算' else 0, axis=1)
df_费用结算数据['调剂金额_新'] = df_费用结算数据.apply(lambda row: -row['金额'] if row['结算状态'] == '结算' else 0, axis=1)

# 按照报账流程类型(工单级别)、渠道编码、分组，对结算金额进行汇总
df_费用结算数据_汇总 = df_费用结算数据.groupby(['报账流程类型(工单级别)', '渠道编码', '渠道名称'])['结算金额'].sum().reset_index()

# 使用ExcelWriter追加模式写入，避免覆盖原有工作表
with pd.ExcelWriter(file_path_费用结算数据稽核, engine='openpyxl', mode='a') as writer:
    # 先删除已存在的原工作表和汇总结算金额表（如果存在）
    for sheet_name in ['合作费用业务明细查询0', '汇总结算金额']:
        if sheet_name in writer.book.sheetnames:
            del writer.book[sheet_name]
    
    # 写入更新后的原工作表数据
    df_费用结算数据.to_excel(writer, index=False, sheet_name='合作费用业务明细查询0')
    
    # 写入汇总结算金额表
    df_费用结算数据_汇总.to_excel(writer, index=False, sheet_name='汇总结算金额')
    
    # 添加总计行
    total_row = pd.DataFrame({
        '报账流程类型(工单级别)': ['总计'],
        '渠道编码': [''],
        '渠道名称': [''],
        '结算金额': [df_费用结算数据_汇总['结算金额'].sum()]
    })
    total_row.to_excel(writer, index=False, sheet_name='汇总结算金额', startrow=len(df_费用结算数据_汇总)+1, header=False)

logging.info("数据处理完成，已更新合作费用业务明细查询0和汇总结算金额表")
