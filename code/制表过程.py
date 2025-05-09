import os
from numpy.random import f
import pandas as pd
import logging
from datetime import datetime
import numpy as np
#项目路径
date_day = datetime.today().strftime('%Y-%m-%d')
basic_path = rf"C:\project\python_project\AI+市场工具&稽核\AI+市场工具&稽核第一期需求\选项目-制表-得出结论"

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
######降档判断######
def 降档判断处理(df_待调剂明细数据, df_待调剂数据, df_降档):
    """
    执行降档判断处理流程
    参数:
        df_待调剂明细数据: 待调剂明细数据DataFrame
        df_待调剂数据: 待调剂数据DataFrame
        df_降档: 降档报表DataFrame
    返回:
        处理后的df_待调剂明细数据
    """
    # 记录初始数据量
    logging.info(f"降档判断处理 - 输入数据量: {len(df_待调剂明细数据)}")
    
    # 添加降档判断列
    df_待调剂数据_降档判断 = df_待调剂数据.drop_duplicates(
        subset=['统计月份', '政策分期小项编码', '报账流程类型(工单级别)']
    )[['统计月份', '政策分期小项编码', '报账流程类型(工单级别)', '降档判断']]
    
    df_待调剂明细数据 = pd.merge(
        df_待调剂明细数据,
        df_待调剂数据_降档判断,
        on=['统计月份', '政策分期小项编码', '报账流程类型(工单级别)'],
        how='left'
    )

    # 初始化降档列为'非降档'并记录
    df_待调剂明细数据['降档'] = '非降档'
    logging.info(f"降档列初始化完成 - 非降档记录数: {len(df_待调剂明细数据)}")

    # 为不需要降档判断的记录添加结算原因
    不需要降档判断 = df_待调剂明细数据['降档判断'] != '需要'
    df_待调剂明细数据.loc[不需要降档判断, '结算原因'] = '|降档:无需判断降档|'
    logging.info(f"无需降档判断的记录数: {不需要降档判断.sum()}")

    # 对需要降档判断的记录进行降档判断
    需要降档判断 = df_待调剂明细数据['降档判断'] == '需要'
    if 需要降档判断.any():
        # 筛选出关联降档报表【是否当月携入移动】为'否'的数据
        df_降档 = df_降档[df_降档['是否当月携入移动'] == '否']
        logging.info(f"降档报表筛选后记录数: {len(df_降档)}")
        
        # 添加降档匹配标记列
        df_降档_匹配 = df_降档[['用户标识', '业务发展月']].drop_duplicates()
        df_降档_匹配['_降档匹配'] = True
        
        df_待调剂明细数据 = pd.merge(
            df_待调剂明细数据,
            df_降档_匹配,
            on=['用户标识', '业务发展月'],
            how='left'
        )
        降档条件 = (df_待调剂明细数据.loc[需要降档判断, '_降档匹配'] == True)

        # 更新需要降档判断的记录的降档状态
        df_待调剂明细数据.loc[需要降档判断, '降档'] = np.where(
            降档条件,
            '降档',
            '非降档'
        )

        df_待调剂明细数据.loc[需要降档判断 & 降档条件, '结算原因'] += '|降档:关联降档剔除结算|'
        
        # 记录降档判断结果
        logging.info(f"降档记录数: {(df_待调剂明细数据.loc[需要降档判断, '降档'] == '降档').sum()}")
        logging.info(f"非降档记录数: {(df_待调剂明细数据.loc[需要降档判断, '降档'] == '非降档').sum()}")
        
        # 删除临时列
        df_待调剂明细数据.drop(columns=['_降档匹配'], inplace=True)
    
    # 删除降档判断列
    df_待调剂明细数据.drop(columns=['降档判断'], inplace=True)
    
    # 检查是否有空值
    if df_待调剂明细数据['降档'].isna().any():
        logging.warning(f"发现降档列有空值记录数: {df_待调剂明细数据['降档'].isna().sum()}")
        # 确保所有记录都有值
        df_待调剂明细数据['降档'] = df_待调剂明细数据['降档'].fillna('非降档')
    
    logging.info(f"降档判断处理完成 - 输出数据量: {len(df_待调剂明细数据)}")
    return df_待调剂明细数据
######待调剂明细数据表与渠道档案报表比碰######
def 渠道档案判断处理(df_待调剂明细数据, df_渠道档案报表):
    """
    执行渠道档案判断处理流程
    参数:
        df_待调剂明细数据: 待调剂明细数据DataFrame
        df_渠道档案报表: 渠道档案报表DataFrame
    返回:
        处理后的df_待调剂明细数据
    """
    # 记录初始数据量
    logging.info(f"渠道档案判断处理 - 输入数据量: {len(df_待调剂明细数据)}")
    
    # 将业务发展月转换为日期格式
    df_待调剂明细数据['业务发展月'] = pd.to_datetime(
        df_待调剂明细数据['业务发展月'], 
        format='%Y%m', 
        errors='coerce'
    )
    
    # 筛选渠道档案报表【渠道状态】为'正常'和'冻结'的数据
    df_渠道档案报表 = df_渠道档案报表[df_渠道档案报表['渠道状态'].isin(['冻结','正常'])].copy()
    
    # 渠道档案报表日期字段转换
    for column in ['协议起始日', '协议终止日']:
        df_渠道档案报表.loc[:,column] = pd.to_datetime(
            df_渠道档案报表[column], 
            format='%Y%m%d', 
            errors='coerce'
        )
    df_渠道档案报表.loc[:,'发展月结算范围'] = pd.to_datetime(
        df_渠道档案报表['发展月结算范围'],
        format='%Y%m',
        errors='coerce'
    )
    
    # 添加渠道档案相关列
    df_渠道档案_匹配 = df_渠道档案报表[['渠道编码', '发展月结算范围', '协议起始日', '协议终止日', '渠道状态']].drop_duplicates().copy()
    #生成文件
    df_渠道档案_匹配.to_excel(
        os.path.join(project_path, '渠道档案_匹配.xlsx'),
        index=False
    )
    df_渠道档案_匹配 = df_渠道档案_匹配.rename(columns={'渠道状态': '_渠道状态'})  # 重命名以避免冲突
    
    df_待调剂明细数据 = pd.merge(
        df_待调剂明细数据,
        df_渠道档案_匹配,
        on=['渠道编码'],
        how='left',
        indicator='_渠道档案匹配'
    )
    
    # 修改后续所有使用"渠道状态"的地方为"_渠道状态"
    df_待调剂明细数据['解约'] = df_待调剂明细数据['_渠道档案匹配'].map({
        'both': '正常',
        'left_only': '冻结'
    })
    # 为冻结状态添加结算原因
    冻结状态 = df_待调剂明细数据['解约'] == '冻结'
    df_待调剂明细数据.loc[冻结状态, '结算原因'] += '|解约:解约渠道剔除结算|'
    
    # 协议终止日判断
    当前日期 = pd.to_datetime('today') - pd.offsets.MonthBegin(1)
    终止日期条件 = (
        (df_待调剂明细数据['_渠道档案匹配'] == 'both') & 
        df_待调剂明细数据['协议终止日'].notna() & 
        (df_待调剂明细数据['协议终止日'] < 当前日期)
    )
    df_待调剂明细数据.loc[终止日期条件, '结算状态'] = '待稽核'
    df_待调剂明细数据.loc[终止日期条件, '解约'] = '待稽核'
     # 添加具体的日期对比信息
    df_待调剂明细数据.loc[终止日期条件, '结算原因'] += df_待调剂明细数据.loc[终止日期条件].apply(
        lambda x: f"|解约:超出渠道协议终止日(终止日:{x['协议终止日'].strftime('%Y%m%d')},当前:{当前日期.strftime('%Y%m%d')})|",
        axis=1
    )
    
    # 处理发展月结算范围
    发展月结算范围_存在记录 = (df_待调剂明细数据['_渠道档案匹配'] == 'both') & df_待调剂明细数据['发展月结算范围'].notna()
    
    if 发展月结算范围_存在记录.any():
        
        # 设置解约状态
        df_待调剂明细数据.loc[发展月结算范围_存在记录, '解约'] = '冻结'
        
        # 正常记录条件
        正常记录条件 = 发展月结算范围_存在记录 & (
            df_待调剂明细数据['业务发展月'] <= df_待调剂明细数据['发展月结算范围']
        )
        df_待调剂明细数据.loc[正常记录条件, '解约'] = '正常'
        
        # 更新结算原因
        冻结记录条件 = 发展月结算范围_存在记录 & (df_待调剂明细数据['解约'] == '冻结')
        df_待调剂明细数据.loc[冻结记录条件, '结算原因'] += '|解约:超出发展月结算范围|'
        df_待调剂明细数据.loc[冻结记录条件, '结算状态'] = '不结算'
    
    # # 删除临时列
    # df_待调剂明细数据.drop(
    #     columns=['_渠道档案匹配', '发展月结算范围', '协议起始日', '协议终止日', '_渠道状态'],
    #     inplace=True
    # )
    
    logging.info(f"渠道档案判断处理完成 - 输出数据量: {len(df_待调剂明细数据)}")
    return df_待调剂明细数据

######待调剂明细数据表异常判断######
def 判重处理(df_待调剂明细数据, df_待调剂数据):
    """
    执行判重处理流程
    参数:
        df_待调剂明细数据: 待调剂明细数据DataFrame
        df_待调剂数据: 待调剂数据DataFrame
    返回:
        处理后的df_待调剂明细数据
    """
    判重指令 = {
        '无需判断':{
            '关联字段':[]
        },
        '用户号码':{
            '关联字段':['用户号码','政策分期小项编码']
        },
        '订购流水':{
            '关联字段':['订购流水','政策分期小项编码']
        },
        '账号ID':{
            '关联字段':['账号id','政策分期小项编码']
        },
        '营销案':{
            '关联字段':['用户号码','政策分期小项编码','业务编码','业务发展月']
        }
    }
    
    # 待调剂明细数据表与待调剂数据比碰
    merged_待调剂明细数据_待调剂数据 = pd.merge(
        df_待调剂明细数据,
        df_待调剂数据[['统计月份','政策分期小项编码','报账流程类型(工单级别)','判重指令']],
        on=['统计月份','政策分期小项编码','报账流程类型(工单级别)'],
        how='left',
        indicator=True
    )

    # 初始化用户重复次数列
    df_待调剂明细数据['用户重复次数'] = 1

    # 对每种判重指令进行处理
    for 指令, 配置 in 判重指令.items():
        if 指令 != '无需判断' and 配置['关联字段']:
            # 获取当前判重指令的记录
            当前指令记录 = merged_待调剂明细数据_待调剂数据['判重指令'] == 指令
            logging.info(f"开始处理判重指令: {指令}, 记录数: {当前指令记录.sum()}")
            if 当前指令记录.any():
                # 按关联字段分组计数
                重复计数 = df_待调剂明细数据.loc[当前指令记录].groupby(配置['关联字段']).size()
                # 更新用户重复次数
                for idx in df_待调剂明细数据.loc[当前指令记录].index:
                    分组值 = tuple(df_待调剂明细数据.loc[idx, field] for field in 配置['关联字段'])
                    重复次数 = 重复计数[分组值]
                    df_待调剂明细数据.loc[idx, '用户重复次数'] = 重复次数
                    df_待调剂明细数据.loc[idx, '判重指令'] = 指令
                    if 重复次数 > 1:
                        if df_待调剂明细数据.loc[idx, '结算状态'] == '结算':
                            df_待调剂明细数据.loc[idx, '结算状态'] = '异常'
                        df_待调剂明细数据.loc[idx, '结算原因'] += f'|判重:{指令}|'
    return df_待调剂明细数据

def 结算状态(df_待调剂明细数据):
    """
    执行结算状态判断
    参数:
        df_待调剂明细数据: 待调剂明细数据DataFrame
    返回:
        处理后的df_待调剂明细数据
    """
    # 筛选结算状态为空的数据
    df_待调剂明细数据_结算状态为空 = df_待调剂明细数据[df_待调剂明细数据['结算状态'].isna()].copy()
    
    # 设置默认值为'不结算'
    df_待调剂明细数据_结算状态为空.loc[:, '结算状态'] = '不结算'
    
    # 降档列为'非降档'与解约列为'正常'的数据结算结果为结算
    df_待调剂明细数据_结算状态为空.loc[
        (df_待调剂明细数据_结算状态为空['降档'] == '非降档') &
        (df_待调剂明细数据_结算状态为空['解约'] == '正常'),
        '结算状态'
    ] = '结算'
    
    # 更新结算状态
    df_待调剂明细数据.loc[df_待调剂明细数据['结算状态'].isna(), '结算状态'] = df_待调剂明细数据_结算状态为空['结算状态']
    return df_待调剂明细数据
######金额处理######
def 金额处理(df_待调剂明细数据):
    """
    执行金额计算流程
    参数:
        df_待调剂明细数据: 待调剂明细数据DataFrame
    返回:
        处理后的df_待调剂明细数据
    """
    # 根据结算状态设置结算金额和调剂金额
    df_待调剂明细数据['结算金额'] = df_待调剂明细数据.apply(lambda x: x['金额'] if x['结算状态'] == '结算' else 0, axis=1)
    df_待调剂明细数据['调剂金额'] = df_待调剂明细数据.apply(lambda x: 0 if x['结算状态'] == '结算' else -x['金额'], axis=1)
    
    return df_待调剂明细数据

def 保存处理结果(df_待调剂明细数据, file_path_费用结算数据稽核):
    """
    保存处理结果到Excel文件
    参数:
        df_待调剂明细数据: 待调剂明细数据DataFrame
        file_path_费用结算数据稽核: 费用结算数据稽核文件路径
    """
    # 保存主数据到费用结算数据稽核文件
    df_待调剂明细数据.to_excel(file_path_费用结算数据稽核, index=False, sheet_name='合作费用业务明细查询0')
    
    # 提取并保存营销案重复用户号码
    df_重复用户号码 = df_待调剂明细数据.loc[
        (df_待调剂明细数据['判重指令'] == '营销案') & 
        (df_待调剂明细数据['用户重复次数'] > 1) &
        (df_待调剂明细数据['结算状态'] == '异常'), 
        '用户号码'
    ].drop_duplicates().to_frame()
    
    with pd.ExcelWriter(file_path_费用结算数据稽核, mode='a', engine='openpyxl') as writer:
        df_重复用户号码.to_excel(writer, sheet_name='营销案重复用户号码', index=False)

######主程序######
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

file_path_待调剂明细数据 = rf"{project_path}\待调剂明细数据.xlsx"
file_path_待调剂数据 = rf"{project_path}\待调剂数据.xlsx"
file_path_降档报表 = rf"{project_path}\降档报表.xlsx"
file_path_渠道档案报表 = rf"{project_path}\渠道档案报表.xlsx"
file_path_费用结算数据稽核 = rf"{project_path}\关于{year_suffix}年{month}月费用结算数据稽核.xlsx"
file_path_制表取数表 = rf"{project_path}\制表取数表.xlsx"
logging.info("文件初始化完成")
try:
    #获取数据表
    logging.info("开始读取数据表...")
    df_待调剂明细数据 = pd.read_excel(file_path_待调剂明细数据)
    df_待调剂明细数据['业务发展月'] = df_待调剂明细数据['业务发展月'].astype(str)
    df_待调剂数据 = pd.read_excel(file_path_待调剂数据, skiprows=1)
    df_待调剂数据 = df_待调剂数据.rename(columns={'政策分期小项':'政策分期小项编码'}, inplace=False)
    df_降档 = pd.read_excel(file_path_降档报表, skiprows=1)
    #df_降档的业务发展月列取前六位
    df_降档['业务发展月'] = df_降档['业务发展月'].astype(str).str[:6]
    df_渠道档案报表 = pd.read_excel(file_path_渠道档案报表, skiprows=1)
    logging.info("数据表读取完成")

    #初始化结算原因
    df_待调剂明细数据['结算原因'] = ''
    #筛选待调剂明细数据表金额大于0的数据
    df_待调剂明细数据 = df_待调剂明细数据[df_待调剂明细数据['金额'] > 0]
    logging.info(f"筛选后的数据条数: {len(df_待调剂明细数据)}")

    #降档判断
    logging.info("开始降档判断处理...")
    df_待调剂明细数据 = 降档判断处理(df_待调剂明细数据, df_待调剂数据, df_降档)
    logging.info("降档判断处理完成")

    #待调剂明细数据表与渠道档案报表比碰
    logging.info("开始渠道档案判断处理...")
    df_待调剂明细数据 = 渠道档案判断处理(df_待调剂明细数据, df_渠道档案报表)
    logging.info("渠道档案判断处理完成")

    #结算状态
    logging.info("开始结算状态...")
    df_待调剂明细数据 = 结算状态(df_待调剂明细数据)
    logging.info("结算状态完成")

    #待调剂明细数据表异常判断
    logging.info("开始判重处理...")
    df_待调剂明细数据 = 判重处理(df_待调剂明细数据, df_待调剂数据)
    logging.info("判重处理完成")

    #金额处理
    logging.info("开始金额处理...")
    df_待调剂明细数据 = 金额处理(df_待调剂明细数据)
    logging.info("金额处理完成")

    #生成处理好的xlsx文件
    logging.info("开始保存处理结果...")
    保存处理结果(df_待调剂明细数据, file_path_费用结算数据稽核)
    logging.info("处理结果保存完成")

    logging.info("所有处理完成")
except Exception as e:
    logging.error(f"处理过程中出现错误: {str(e)}", exc_info=True)
    raise



