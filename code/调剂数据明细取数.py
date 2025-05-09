import pandas as pd
import logging
#读取待调剂数据模板，按照调剂月份、报账流程类型(工单级别)分组，整合分组的政策编码，生成列表【‘统计月份’】、【‘报账流程类型(工单级别)’】、【‘政策编码’】

# 读取Excel文件
df = pd.read_excel(file_path_费用结算数据稽核, sheet_name='合作费用业务明细查询0')
df = df[df['调剂金额'] < 0]
# 按照需求分组并整合政策编码
result = df.groupby(['统计月份', '报账流程类型(工单级别)'])['政策编码'].apply(
    lambda x: ','.join(str(code) for code in sorted(set(x)))
).reset_index()
# 添加空文件密码列
result['文件密码'] = ""
# 添加文件路径，使用递增序号
result['文件路径'] = result.index.map(lambda x: rf"{file_path_调剂数据明细取数目录}\调剂数据明细取数表{x+1}.xlsx")
# 输出结果
# for index, row in result.iterrows():
#     统计月份 = row['统计月份']
#     报账类型 = row['报账流程类型(工单级别)']
#     政策编码列表 = row['政策编码']
    # print(f"第{index+1}行: 统计月份={统计月份}, 报账类型={报账类型}, 政策编码={政策编码列表}")
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
    
result.to_excel(file_path_调剂数据明细取数表, index=False)
logging.info(f"\n制表取数表制作完成: {file_path_调剂数据明细取数表}")