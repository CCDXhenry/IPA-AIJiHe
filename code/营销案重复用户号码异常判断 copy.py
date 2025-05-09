import os
import pandas as pd
import logging
#数据表
# 设置日志
logging.info("开始处理数据...")
df_费用结算数据 = pd.read_excel(file_path_费用结算数据稽核, sheet_name='合作费用业务明细查询0')
for index in range(len(user_mobile_paths)):
    user_mobile = user_mobiles[index]
    user_mobile_path = user_mobile_paths[index]
    df_用户号码 = pd.read_excel(user_mobile_path)
    df_费用结算数据_user_mobile = df_费用结算数据[(df_费用结算数据['用户号码'] == user_mobile) & (df_费用结算数据['结算状态'] == '异常')]
    
    # 提取业务发展月并去重
    业务发展月 = df_费用结算数据_user_mobile['业务发展月'].drop_duplicates().copy()
    
    for 业务发展月_value in 业务发展月:
        # 获取该业务发展月的记录数
        费用结算记录数 = len(df_费用结算数据_user_mobile[df_费用结算数据_user_mobile['业务发展月'] == 业务发展月_value])
        
        # 判断用户号码在该业务发展月的登记记录数是否一致
        用户号码记录数 = len(df_用户号码[
            (pd.to_datetime(df_用户号码['登记时间']).dt.to_period('M') == pd.to_datetime(业务发展月_value).to_period('M')) &
            (df_用户号码['取消时间'].isna() | 
             (pd.to_datetime(df_用户号码['取消时间']).dt.to_period('M') > pd.to_datetime(业务发展月_value).to_period('M')))
        ])
        
        # 如果记录数一致，则修改结算状态为正常
        if 费用结算记录数 == 用户号码记录数:
            df_费用结算数据.loc[
                (df_费用结算数据['用户号码'] == user_mobile) & 
                (df_费用结算数据['业务发展月'] == 业务发展月_value) &
                (df_费用结算数据['结算状态'] == '异常'),
                '结算状态'
            ] = '结算'
            logging.info(f"用户号码 {user_mobile} 在业务发展月 {业务发展月_value} 的记录数一致，已修改为正常状态")

# 保存修改后的结果，使用追加模式保留其他工作表
with pd.ExcelWriter(file_path_费用结算数据稽核, engine='openpyxl', mode='a') as writer:
    # 先删除已存在的原工作表（如果存在）
    if '合作费用业务明细查询0' in writer.book.sheetnames:
        idx = writer.book.sheetnames.index('合作费用业务明细查询0')
        writer.book.remove(writer.book.worksheets[idx])
        writer.book.create_sheet('合作费用业务明细查询0', idx)
    
    # 写入修改后的数据
    df_费用结算数据.to_excel(writer, index=False, sheet_name='合作费用业务明细查询0')

logging.info("数据保存完成，已更新合作费用业务明细查询0工作表")
