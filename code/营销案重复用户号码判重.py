import pandas as pd
import logging
from datetime import datetime
basic_path = rf"C:\project\python_project\AI+市场工具&稽核\AI+市场工具&稽核第一期需求\选项目-制表-得出结论"
# 获取当前年月
current_date = datetime.now()
year_suffix = str(current_date.year)[2:]  # 获取年份后两位
month = current_date.month
day = current_date.day
project_path = rf"{basic_path}\{year_suffix}年{month}月{day}日数据"
file_path_费用结算数据稽核 = rf"C:\project\python_project\AI+市场工具&稽核\AI+市场工具&稽核第一期需求\选项目-制表-得出结论\23年10月27日数据\费用结算数据稽核.xlsx"
try:
    df_营销案重复用户号码 = pd.read_excel(file_path_费用结算数据稽核, sheet_name='营销案重复用户号码')
    user_mobiles = df_营销案重复用户号码['用户号码'].tolist()
    user_mobiles = list(set(user_mobiles))
    logging.info(f"读取{file_path_费用结算数据稽核}成功，读取到{len(user_mobiles)}个营销案重复用户号码")
    #以用户号码为名，为每个用户号码创建一个xls文件路径，保存在列表
    user_mobile_paths = [rf"{project_path}\{user_mobile}.xls" for user_mobile in user_mobiles]

except Exception as e:
    logging.error(f"读取{file_path_费用结算数据稽核}失败，错误信息：{str(e)}")
