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
#读取待调剂数据模板，按照调剂月份、报账流程类型(工单级别)分组，整合分组的政策编码，生成列表【‘统计月份’】、【‘报账流程类型(工单级别)’】、【‘政策编码’】
file_path_待调剂数据 = r"C:\project\python_project\AI+市场工具&稽核\AI+市场工具&稽核第一期需求\选项目-制表-得出结论\待调剂数据.xlsx"
# 读取Excel文件
df = pd.read_excel(file_path_待调剂数据, skiprows=1)

# 按照需求分组并整合政策编码
result = df.groupby(['统计月份', '报账流程类型(工单级别)'])['政策编码'].apply(
    lambda x: ','.join(str(code) for code in sorted(set(x)))
).reset_index()
# 添加空文件密码列
result['文件密码'] = ""
# 添加文件路径，使用递增序号
result['文件路径'] = result.index.map(lambda x: rf"{project_path}\制表取数表{x+1}.xlsx")
# 将结果转换为列表
result_list = []
for index, row in result.iterrows():
    result_list.append({
        '统计月份': row['统计月份'],
        '报账流程类型': row['报账流程类型(工单级别)'],
        '政策编码': row['政策编码'],
        '文件密码': row['文件密码'],
        '文件路径': row['文件路径']
    })
print(result_list[0]['统计月份'])
result_list[result_index]['统计月份']
#获取列表长度
length = len(result_list)
print(result_list)
# 可选：将结果保存到新Excel文件
output_path = r"C:\project\python_project\AI+市场工具&稽核\AI+市场工具&稽核第一期需求\选项目-制表-得出结论\待调剂数据结果.xlsx"
result.to_excel(output_path, index=False)
print(f"\n结果已保存到: {output_path}")