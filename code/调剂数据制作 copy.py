import os
from numpy.random import f
import pandas as pd
import logging
from datetime import datetime
import numpy as np

######主程序######
#获取数据表
logging.info("开始读取数据表...")
df_费用结算数据稽核 = pd.read_excel(file_path_费用结算数据稽核, sheet_name='合作费用业务明细查询0')
df_调剂数据导入模板 = pd.read_excel(file_path_调剂数据导入模板, engine='xlrd')
df_调剂依据文件 = pd.read_excel(file_path_调剂依据文件)
logging.info("数据表读取完成")
# 获取依据文件号与依据文件名称
# 调剂依据文件按照‘合作费用政策分期小项编码‘去重
df_调剂依据文件 = df_调剂依据文件.drop_duplicates(subset=['合作费用政策分期小项编码'])
# ‘合作费用政策分期小项编码‘重命名’政策分期小项编码’
df_调剂依据文件 = df_调剂依据文件.rename(columns={'合作费用政策分期小项编码': '政策分期小项编码'})
# ‘文件号’重命名为‘依据文号(必填)’
df_调剂依据文件 = df_调剂依据文件.rename(columns={'文件号': '依据文号(必填)'})
# ‘文件名称’重命名为‘依据文件名称(必填)’
df_调剂依据文件 = df_调剂依据文件.rename(columns={'文件名称': '依据文件名称(必填)'})

# 获取费用结算数据稽核’调剂金额‘<0的数据
df_费用结算数据稽核 = df_费用结算数据稽核[df_费用结算数据稽核['调剂金额'] < 0]
# merge费用结算数据稽核与调剂依据文件，获取依据文号(必填)与依据文件名称(必填)
print(f"合并前数据量 - 费用结算数据: {len(df_费用结算数据稽核)}")
print(f"合并前数据量 - 调剂依据文件: {len(df_调剂依据文件)}")

df_费用结算数据稽核 = pd.merge(
    df_费用结算数据稽核, 
    df_调剂依据文件[['政策分期小项编码', '依据文号(必填)', '依据文件名称(必填)']],
    on='政策分期小项编码',
    how='left'
    )

print(f"合并后数据量: {len(df_费用结算数据稽核)}")

# 费用结算数据稽核按照‘薪酬类型’分类，‘代办渠道费用’、’委托加盟营业厅’为合作费类
df_费用结算数据稽核_合作费 = df_费用结算数据稽核[
    (df_费用结算数据稽核['薪酬类型'] == '代办渠道费用') | 
    (df_费用结算数据稽核['薪酬类型'] == '委托加盟营业厅')
].copy()

# 营销支撑服务费管理为营销类
df_费用结算数据稽核_营销类 = df_费用结算数据稽核[
    df_费用结算数据稽核['薪酬类型'] == '营销支撑服务费管理'
].copy()

# 在上面的基础上按照'报账流程类型(工单级别)'分类，市级报账(市级工单)为市级类，县级报账(县级工单)为县级类

# 合作费类按工单级别分类
df_费用结算数据稽核_合作费_市级 = df_费用结算数据稽核_合作费[
    df_费用结算数据稽核_合作费['报账流程类型(工单级别)'] == '市级报账(市级工单)'
].copy()

df_费用结算数据稽核_合作费_县级 = df_费用结算数据稽核_合作费[
    df_费用结算数据稽核_合作费['报账流程类型(工单级别)'] == '县级报账(县级工单)'
].copy()

# 营销类按工单级别分类
df_费用结算数据稽核_营销类_市级 = df_费用结算数据稽核_营销类[
    df_费用结算数据稽核_营销类['报账流程类型(工单级别)'] == '市级报账(市级工单)'
].copy()

df_费用结算数据稽核_营销类_县级 = df_费用结算数据稽核_营销类[
    df_费用结算数据稽核_营销类['报账流程类型(工单级别)'] == '县级报账(县级工单)'
].copy()

# 按照‘统计月份‘、‘政策分期小项编码‘分组，根据调剂数据导入模板新建保存数据，文件名为’薪酬类型‘-’工单级别‘-‘统计月份‘-’政策分期小项编码‘-‘渠道归属县市’.xls
'''
# 调剂数据导入模对应列-费用结算数据稽核对应列：
# 业务发展月份(必填;格式:YYYYMM) - 统计月份，
# 渠道编码(必填;数值型;判断:已存在的渠道) - 渠道编码， 
# 结算规则编码(数值型:必填,注意：这里是结算规则编码，不是合作费用规则编码，其中合作费用政策分期小项编码、合作费用规则编码、结算规则编码 三者对应关系，请到“菜单路径：合作费用管理-合作费用配置管理-合作费用规则查询视图”中查询。)-结算规则编码，
# 用户号码(数值型，必填) - 用户号码
# 调剂金额（必填，数值型，允许有两位小数点,单位：元。调剂金额格式仅支持常规格式(如1231.03),不支持货币(￥1,313.12)、千分位(如1,313.12)、文本(如1231.03元)、特殊字符、会计专用、科学计数、自定义等格式） - 调剂金额,
# 调剂原因(必填)（如果是结算数据源导入的合作费用政策，必须要把不计酬的结算数据源记录导入，调剂金额填0，把无效原因写在调剂原因这列，用于在IOP、销售宝的合作费用对账单明细中展示不计酬原因，这是合作费用透明化展示的要求！请合作费用调剂人员一定要找市公司合作费用政策负责人要无效清单！） - 结算原因
# 调剂类型(必填)稽核补发、延迟发放、政策性补发、其它补发、稽核扣回、实名制违规扣罚、终端违规扣罚、养卡套利扣罚、其他违规扣罚 - 默认'5.稽核扣回'
# 依据文号(必填) - 依据文号(必填)
# 依据文件名称(必填) - 依据文件名称(必填)
'''
# 定义处理函数，按照分组导出文件
def export_grouped_data(df):
    # 按照统计月份和政策分期小项编码分组
    grouped = df.groupby(['统计月份', '政策分期小项编码'])
    
    # 保存所有生成的文件信息
    file_info_list = []
    
    # 查找对应的列名
    结算规则编码列 = find_column_name(df_调剂数据导入模板, '结算规则编码')
    调剂金额列 = find_column_name(df_调剂数据导入模板, '调剂金额')
    调剂原因列 = find_column_name(df_调剂数据导入模板, '调剂原因')
    调剂类型列 = find_column_name(df_调剂数据导入模板, '调剂类型')

    for (统计月份, 政策分期小项编码), group in grouped:
        # 获取薪酬类型和工单级别(从第一行获取)
        薪酬类型 = group['薪酬类型'].iloc[0]
        工单级别 = group['报账流程类型(工单级别)'].iloc[0]
        
        if 工单级别 == '县级报账(县级工单)':
            # 县级数据需要额外按照渠道归属县市分组
            county_grouped = group.groupby('渠道/用户归属县市')
            for 渠道归属县市, county_group in county_grouped:
                # 创建输出DataFrame
                output_df = pd.DataFrame(columns=df_调剂数据导入模板.columns)
                # 映射字段
                output_df['业务发展月份(必填;格式:YYYYMM)'] = county_group['统计月份']
                output_df['渠道编码(必填;数值型;判断:已存在的渠道)'] = county_group['渠道编码']
                output_df[结算规则编码列] = county_group['结算规则编码']
                output_df['用户号码(数值型，必填)'] = county_group['用户号码']
                output_df[调剂金额列] = county_group['调剂金额']
                output_df[调剂原因列] = county_group['结算原因']
                output_df[调剂类型列] = '5.稽核扣回'
                output_df['依据文号(必填)'] = county_group['依据文号(必填)']
                output_df['依据文件名称(必填)'] = county_group['依据文件名称(必填)']
                
                # 生成文件名
                filename = f"{薪酬类型}-{工单级别}-{统计月份}-{政策分期小项编码}-{渠道归属县市}.xlsx"
                filepath = os.path.join(file_path_调剂表数据目录, filename)
                print(f"文件名: {filename}")
                # 保存文件信息
                file_info = {
                    '薪酬类型': 薪酬类型,
                    '工单级别': 工单级别,
                    '统计月份': 统计月份,
                    '政策分期小项编码': 政策分期小项编码,
                    '渠道归属县市': 渠道归属县市,
                    '文件路径': filepath
                }
                file_info_list.append(file_info)
                
                # 保存文件
                output_df.to_excel(filepath, index=False, engine='openpyxl')
                logging.info(f"已生成县级文件: {filename}")
        else:
            # 市级数据处理
            output_df = pd.DataFrame(columns=df_调剂数据导入模板.columns)
            output_df['业务发展月份(必填;格式:YYYYMM)'] = group['统计月份']
            output_df['渠道编码(必填;数值型;判断:已存在的渠道)'] = group['渠道编码']
            output_df[结算规则编码列] = group['结算规则编码']
            output_df['用户号码(数值型，必填)'] = group['用户号码']
            output_df[调剂金额列] = group['调剂金额']
            output_df[调剂原因列] = group['结算原因']
            output_df[调剂类型列] = '5.稽核扣回'
            output_df['依据文号(必填)'] = group['依据文号(必填)']
            output_df['依据文件名称(必填)'] = group['依据文件名称(必填)']
            
            filename = f"{薪酬类型}-{工单级别}-{统计月份}-{政策分期小项编码}.xlsx"
            filepath = os.path.join(file_path_调剂表数据目录, filename)
            print(f"文件名: {filename}")
            # 保存文件信息
            file_info = {
                '薪酬类型': 薪酬类型,
                '工单级别': 工单级别,
                '统计月份': 统计月份,
                '政策分期小项编码': 政策分期小项编码,
                '渠道归属县市': None,
                '文件路径': filepath
            }
            file_info_list.append(file_info)
            
            output_df.to_excel(filepath, index=False, engine='openpyxl')
            logging.info(f"已生成市级文件: {filename}")
    
    return file_info_list

def find_column_name(template_df, keyword):
    """从模板中查找包含关键字的列名"""
    for col in template_df.columns:
        if keyword in col:
            return col
    raise ValueError(f"未找到包含'{keyword}'的列名")

# 处理各类数据并保存文件信息
all_file_info = []
all_file_info.extend(export_grouped_data(df_费用结算数据稽核_合作费_市级))
all_file_info.extend(export_grouped_data(df_费用结算数据稽核_合作费_县级))
all_file_info.extend(export_grouped_data(df_费用结算数据稽核_营销类_市级))
all_file_info.extend(export_grouped_data(df_费用结算数据稽核_营销类_县级))

# 保存所有文件信息到CSV文件
file_info_df = pd.DataFrame(all_file_info)
file_info_df.to_csv(file_path_生成文件信息表, index=False, encoding='utf-8-sig')