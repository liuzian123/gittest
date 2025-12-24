import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta
#定义地址
#os.chdir(r"C:\Users\45531\Desktop\ebay核算数据及代码\2025-6\辅助数据")

# 获取上一个月和对应年份
#prev_year, prev_month = (datetime.today().year - 1, 12) if datetime.today().month == 1 else (datetime.today().year,datetime.today().month - 1)
#汇率 = f"{(datetime.now().replace(day=1) - timedelta(days=1)).strftime('%Y.%m')}月汇率"
# prev_year, prev_month =2025,5
# 汇率 = f"2025.05月汇率"
#print(汇率)

# 设置后台数据所在目录
#auxiliary_data_path = r"C:\Users\45531\Desktop\ebay核算数据及代码\2025-6\辅助数据"
#backend_data_path   = r"C:\Users\45531\Desktop\ebay核算数据及代码\2025-6\后台数据"
#result_path         = r"C:\Users\45531\Desktop\ebay核算数据及代码\2025-6"


# ===== 用户自定义变量 =====
project_path = r"C:\Users\45531\Desktop\ebay核算数据及代码"  # 你可以在这里修改项目地址
# =========================

# 切换到项目目录
os.chdir(project_path)

# 获取上一个月和对应年份
prev_year, prev_month = (datetime.today().year - 1, 12) if datetime.today().month == 1 else (datetime.today().year, datetime.today().month - 1)
汇率 = f"{(datetime.now().replace(day=1) - timedelta(days=1)).strftime('%Y.%m')}月汇率"
# prev_year, prev_month = 2025, 5  # 手动调试时可以取消注释
# 汇率 = f"2025.05月汇率"          # 手动调试时可以取消注释
print(汇率)

# 设置后台数据所在目录
backend_data_path = os.path.join(f"{prev_year}-{prev_month}", "后台数据")
auxiliary_data_path = os.path.join(f"{prev_year}-{prev_month}", "辅助数据")
result_path = os.path.normpath(f"{prev_year}-{prev_month}")


print("辅助数据路径：", os.path.abspath(auxiliary_data_path))
print("目录内容：", os.listdir(auxiliary_data_path))

lingxing_order_data_initial = pd.read_excel(os.path.join(auxiliary_data_path,'领星导出-订单管理.xlsx'), dtype={"ASIN/商品Id":str,"系统单号":str,"订单商品ID":str,"参考号":str})
lingxing_order_data_initial=lingxing_order_data_initial[lingxing_order_data_initial["平台"] == "eBay"]


full_path = os.path.abspath(os.path.join(auxiliary_data_path,'领星导出-订单管理.xlsx'))
print("完整路径：", full_path)
print("文件存在？", os.path.exists(full_path))




# 获取上一个月和对应年份
prev_year, prev_month = (datetime.today().year - 1, 12) if datetime.today().month == 1 else (datetime.today().year,datetime.today().month - 1)
汇率 = f"{(datetime.now().replace(day=1) - timedelta(days=1)).strftime('%Y.%m')}月汇率"

rate_data=pd.read_excel(os.path.join(auxiliary_data_path,"最新负责人-确认版.xlsx"),sheet_name="汇率")

# 英文列名到中文列名的映射字典
en_to_cn_column_mapping = {
    "Transaction creation date": "交易创建日期",
    "Type": "类型",
    "Order number": "订单编号",
    "Legacy order ID": "旧订单编号",
    "Buyer username": "买家用户名",
    "Buyer name": "买家姓名",
    "Ship to city": "收货人所在县/市",
    "Ship to province/region/state": "运送至省/地区/州",
    "Ship to zip": "收货人邮政编码",
    "Ship to country": "收货人所在国家/地区",
    "Net amount": "净额",
    "Payout currency": "发款货币",
    "Payout date": "发款日期",
    "Payout ID": "发款编号",
    "Split payout ID": "拆分发款编号",
    "Payout method": "收款方式",
    "Payout status": "发款状态",
    "Reason for hold": "冻结原因",
    "Item ID": "物品编号",
    "Transaction ID": "交易编号",
    "Item title": "物品标题",
    "Custom label": "自定义标签",
    "Quantity": "数量",
    "Item subtotal": "物品小计",
    "Shipping and handling": "运费与处理费",
    "Seller collected tax": "卖家收取的税费",
    "eBay collected tax": "eBay 收取的税费",
    "Seller specified VAT rate": "卖家指定的增值税税率",
    "Final Value Fee - fixed": "成交费 — 固定",
    "Final Value Fee - variable": "成交费 — 因品类而异",
    "Regulatory operating fee": "监管运营费",
    "Very high \"item not as described\" fee": "“物品与描述不符”指数非常高的费用",
    "Below standard performance fee": "表现不合格的费用",
    "International fee": "跨国交易费用",
    "Gross transaction amount": "交易总金额",
    "Transaction currency": "交易货币",
    "Exchange rate": "汇率",
    "Reference ID": "参考编号",
    "Description": "描述"
}





#读取费用类型数据
fee_type_path = os.path.join(auxiliary_data_path, '费用类型+头程.xlsx')
fee_type_df = pd.read_excel(fee_type_path,sheet_name="费用类型")
fee_type_df = fee_type_df.fillna("")

# 定义要匹配的文件名关键词
file_pattern = "-交易报告"

# 查找所有包含该关键词的 CSV 文件
matching_files = [f for f in os.listdir(backend_data_path) if file_pattern in f]

# 用于存储所有处理后的数据
all_data = []

# 遍历每个匹配到的文件
for file in matching_files:
    file_path = os.path.join(backend_data_path, file)

    # 获取店铺名称（第9行第2列）
    shop_name = pd.read_csv(file_path, header=None, nrows=9).iloc[8, 1]

    # 判断列名是英文还是中文，并统一映射为中文
    df_header_check = pd.read_csv(file_path, nrows=0, header=11)
    original_columns = df_header_check.columns.tolist()

    if any(col in en_to_cn_column_mapping.keys() for col in original_columns):
        current_data = pd.read_csv(file_path, header=11)
        current_data.rename(columns=en_to_cn_column_mapping, inplace=True)
    else:
        current_data = pd.read_csv(file_path, header=11)

    current_data['账号名称'] = shop_name.lower()
    all_data.append(current_data)

# 将所有文件的数据合并
merged_order_data = pd.concat(all_data, ignore_index=True)

#读取 eBay 账号信息表, 创建账号名称到新领星代码的映射字典，防止大小写干扰店铺匹配
account_info_path = os.path.join(backend_data_path,'EBAY账号信息表.xlsx')
account_info_df = pd.read_excel(account_info_path)
account_info_df['eBay账户']=account_info_df['eBay账户'].astype(str).str.lower()
account_info_df['账号名称']=account_info_df['账号名称'].astype(str).str.lower()

# 先尝试基于 eBay账户 的映射
account_mapping_by_ebay = dict(zip(account_info_df['eBay账户'], account_info_df['新领星代码']))
merged_order_data['店铺'] = merged_order_data['账号名称'].map(account_mapping_by_ebay)

# 找出店铺未匹配成功的行
unmatched_mask = merged_order_data['店铺'].isna() | (merged_order_data['店铺'] == "")

# 基于账号名称的二次映射
account_mapping_by_name = dict(zip(account_info_df['账号名称'], account_info_df['新领星代码']))

# 对未匹配的行进行二次映射
merged_order_data.loc[unmatched_mask, '店铺'] = merged_order_data.loc[unmatched_mask, '账号名称'].map(account_mapping_by_name)

# 将店铺列挪到第一列
shop_column = merged_order_data.pop('店铺')
merged_order_data.insert(0, '店铺', shop_column)

#将整个merged_order_data中的--替换为""
merged_order_data.replace("--", "", inplace=True)

# 清除 "描述" 列中的前后空格
merged_order_data['描述'] = merged_order_data['描述'].str.strip()

#新增"账单月份"列
merged_order_data["账单月份"] = f"{prev_year}-{prev_month:02d}"

# 将需要计算的列转为数值类型，空值填充为0
target_cols = ["成交费 — 固定", "成交费 — 因品类而异", "监管运营费","“物品与描述不符”指数非常高的费用", "表现不合格的费用", "跨国交易费用","慈善捐款", "订金处理费"]

#确保"慈善捐款", "订金处理费"不存在也不会报错
existing_cols = [col for col in target_cols if col in merged_order_data.columns]
merged_order_data[existing_cols] = merged_order_data[existing_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

#新增"交易佣金"列
merged_order_data["交易佣金"] = merged_order_data[existing_cols].sum(axis=1)

#新增"汇率"列
currency_to_rate = rate_data.drop_duplicates(subset=['币种二字码'], keep='first').set_index('币种二字码')[汇率].to_dict()
merged_order_data["汇率2"] = merged_order_data["交易货币"].map(currency_to_rate)

#汇率和交易总金额转为数值，空值填充0
merged_order_data[["交易总金额","交易佣金"]] = merged_order_data[["交易总金额","交易佣金"]].apply(pd.to_numeric, errors='coerce').fillna(0)

# 新增“交易总金额（RMB）”列
merged_order_data["交易总金额（RMB）"] = np.where(
    merged_order_data["交易佣金"] != 0,
    merged_order_data["交易佣金"] * merged_order_data["汇率2"],
    merged_order_data["交易总金额"] * merged_order_data["汇率2"]
)

# 确保类型表中的列名是英文的（根据您提供的column_mapping）
fee_type_df.columns = ['Type', 'Description', '大类', '类型1', '类型2', '类型3', '费用标签类型','备注']

# 重置索引以避免性能问题
merged_order_data.reset_index(drop=True, inplace=True)

# 创建 fee_type_df 的映射表
type_map = fee_type_df[['Type', 'Description', '大类', '类型1', '类型2', '类型3', '费用标签类型']].copy()

# 保留每个 Type 的第一条记录（避免重复映射）
type_map_unique = type_map.drop_duplicates(subset=['Type'], keep='first')

# 类型2：基于 Type + Description -> 类型2
type_desc_to_type2 = dict(zip(type_map['Type'] + type_map['Description'], type_map['类型2']))
merged_order_data['类型_描述'] = merged_order_data['类型'] + merged_order_data['描述']
merged_order_data['类型2'] = merged_order_data['类型_描述'].map(type_desc_to_type2).fillna('')

# 类型1：基于 Type -> 类型1
type_to_type1 = dict(zip(type_map_unique['Type'], type_map_unique['类型1']))
merged_order_data['类型1'] = merged_order_data['类型'].map(type_to_type1).fillna('')

# 大类：基于 Type -> 大类
type_to_category = dict(zip(type_map_unique['Type'], type_map_unique['大类']))
merged_order_data['大类'] = merged_order_data['类型'].map(type_to_category).fillna('')

# 如果 核算科目 为空，则尝试使用 描述 字段去匹配类型表中的 类型2，只要描述中包含映射字典的键（即 Description），就将对应的 类型2 赋值给 核算科目。
# 创建 Description -> 类型2 的映射字典
# 创建 Description -> (类型2, 类型1) 的映射字典
desc_to_type2_and_type1 = dict(zip(fee_type_df['Description'],
                                  zip(fee_type_df['类型2'], fee_type_df['类型1'])))

# 定义一个函数，用于模糊匹配描述并返回第一个匹配到的类型2（同时检查类型1是否匹配）
# 修改模糊匹配函数，改为前10个字符匹配
def match_type2_from_desc_with_type1_check(row, mapping):
    desc = str(row['描述']) if not pd.isna(row['描述']) else ''
    current_type1 = str(row['类型1']) if not pd.isna(row['类型1']) else ''
    if not desc:
        return ''

    # 只取前10个字符进行匹配
    desc_prefix = desc[:10]

    # 先尝试精确匹配前10个字符
    for key in mapping:
        key_prefix = key[:10]
        if key_prefix in desc_prefix and key_prefix != '':
            mapped_type2, mapped_type1 = mapping[key]
            if mapped_type1 == current_type1:
                return mapped_type2

    return ''


# 找出类型2为空的行
empty_type2_mask = merged_order_data['类型2'].isna() | (merged_order_data['类型2'] == '')

# 如果类型2为空，尝试通过描述模糊匹配类型2
if empty_type2_mask.any():
    merged_order_data.loc[empty_type2_mask, '类型2'] = \
        merged_order_data.loc[empty_type2_mask].apply(match_type2_from_desc_with_type1_check,
                                                     mapping=desc_to_type2_and_type1, axis=1)

# 现在基于更新后的类型2重新计算核算科目和费用标签类型
# 核算科目：基于 大类+类型1+类型2 -> 类型3
category_type1_type2_to_type3 = dict(
    zip(type_map['大类'] + type_map['类型1'] + type_map['类型2'], type_map['类型3'])
)
merged_order_data['大类_类型1_类型2'] = merged_order_data['大类'] + merged_order_data['类型1'] + merged_order_data['类型2']
merged_order_data['核算科目'] = merged_order_data['大类_类型1_类型2'].map(category_type1_type2_to_type3).fillna('')

# 费用标签类型：基于 类型3 -> 费用标签类型
type3_to_status = dict(zip(type_map['类型3'], type_map['费用标签类型']))
merged_order_data['费用标签类型'] = merged_order_data['核算科目'].map(type3_to_status).fillna('')

# 删除临时列
merged_order_data.drop(['类型_描述', '大类_类型1_类型2'], axis=1, inplace=True)

# 准备最终输出数据 - 确保所有列都存在并按正确顺序排列
required_columns = ['店铺'] + list(en_to_cn_column_mapping.values()) + [
    '账号名称', '账单月份','汇率2','交易佣金','核算科目' , '费用标签类型','大类', '类型1', '类型2' , '交易总金额（RMB）'
]

# 创建包含所有列的DataFrame
final_merged_data = pd.DataFrame(columns=required_columns)
# 将订单编号转换为字符串，防止后面的额外处理汇总金额失效
final_merged_data["订单编号"]=merged_order_data["订单编号"].astype(str)

# 填充数据
for col in final_merged_data.columns:
    if col in merged_order_data.columns:
        final_merged_data[col] = merged_order_data[col]
    else:
        final_merged_data[col] = ''

# 确保列顺序正确
final_merged_data = final_merged_data[required_columns]

# 输出结果到Excel文件
output_path = os.path.join(result_path, f"后台数据整理-{prev_year}年{prev_month}月.xlsx")
print(os.path.join(result_path, f"后台数据整理-{prev_year}年{prev_month}月.xlsx"))
final_merged_data.to_excel(output_path, engine='xlsxwriter', index=False)

# 筛选"费用标签类型"为"按店铺销售额占比分摊"的数据
filtered_data = final_merged_data[final_merged_data['费用标签类型'] == '按店铺销售额占比分摊'].copy()

# 筛选"核算科目"为"广告费分摊"的数据，保存为独立Excel文件
ad_cost_data = filtered_data[filtered_data['核算科目'] == '广告费分摊']
if not ad_cost_data.empty:
    ad_output_path = os.path.join(result_path, "广告费分摊-按店铺销售额分摊.xlsx")
    ad_cost_data.to_excel(ad_output_path, index=False, engine='xlsxwriter')

# 筛选"核算科目"为"平台其他支出分摊"的数据，保存为独立Excel文件
other_cost_data = filtered_data[filtered_data['核算科目'] == '平台其他支出分摊']
if not other_cost_data.empty:
    other_output_path = os.path.join(result_path, "平台其他支出分摊-按店铺销售额分摊.xlsx")
    other_cost_data.to_excel(other_output_path, index=False, engine='xlsxwriter')

lingxing_order_data_initial = pd.read_excel(os.path.join(auxiliary_data_path,'领星导出-订单管理.xlsx'), dtype={"ASIN/商品Id":str,"系统单号":str,"订单商品ID":str,"参考号":str})
lingxing_order_data_initial=lingxing_order_data_initial[lingxing_order_data_initial["平台"] == "eBay"]


# 筛选费用标签类型：“需额外整理”且核算科目为“广告费”，复制到一个新表，表格命名为广告费-需额外整理
ads_fee = final_merged_data[(final_merged_data['费用标签类型'] == '需额外整理') & (final_merged_data['核算科目'] == '广告费')]
# 复制广告费-需额外整理中的订单编号去领星-订单管理-批量搜索订单编号--不合并单元格导出-导出字段全选（不勾图片）--导出，导出来数据放到领星对应订单明细里面
lingxing_order_data=lingxing_order_data_initial.copy()

# 获取所有在 ads_fee 中存在的订单编号（即广告费订单）
ad_order_numbers = ads_fee['订单编号'].unique()
# 在 lingxing_order_data 中筛选出平台单号存在于 ad_order_numbers 的行
lingxing_order_data = lingxing_order_data[lingxing_order_data['平台单号'].isin(ad_order_numbers)]
print(len(ad_order_numbers),len(lingxing_order_data))

#确保系统单号和商品ID是字符串类型，免得被科学计数法了
lingxing_order_data['系统单号'] = lingxing_order_data['系统单号'].fillna('').astype(str)
lingxing_order_data['订单商品ID'] = lingxing_order_data['订单商品ID'].fillna('').astype(str)
lingxing_order_data['ASIN/商品Id'] = lingxing_order_data['ASIN/商品Id'].fillna('').astype(str)
lingxing_order_data['参考号'] = lingxing_order_data['参考号'].fillna('').astype(str)

# 删除 SKU 为 NaN 或 空字符串 的行
lingxing_order_data = lingxing_order_data[lingxing_order_data['SKU'].notna() & (lingxing_order_data['SKU'] != '')]

#创建一个平台单号到广告费的映射字典
ad_fee_sum_dict = ads_fee.groupby('订单编号')['交易总金额（RMB）'].sum().to_dict()
lingxing_order_data['广告费汇总'] = lingxing_order_data['平台单号'].map(ad_fee_sum_dict)

#读取产品信息表（用于映射产品采购单价）
product_file_path = os.path.join(auxiliary_data_path,r"领星+店小秘产品信息表.xlsx")
product_information = pd.read_excel(product_file_path, sheet_name='普通+组合产品',header=1)

#获取 SKU 对应的「单个产品成本价格」，用于计算「单个产品成本价格」,构建 SKU 到「采购单价（核算）」的映射字典
sku_to_cost_price = product_information.set_index('*SKU')['采购单价（核算）'].to_dict()
lingxing_order_data['单个产品成本价格'] = lingxing_order_data['SKU'].map(sku_to_cost_price)

# 使用 map 添加「单个产品成本价格」列,为每个订单商品匹配其对应的产品成本价，用于后续成本核算
lingxing_order_data['单个订单成本'] = lingxing_order_data['数量'] * lingxing_order_data['单个产品成本价格']

#单个订单成本 = 数量 * 单个产品成本价格,用于统计每个商品在订单中的总成本
order_total_cost = lingxing_order_data.groupby('平台单号')['单个订单成本'].sum().to_dict()
lingxing_order_data['订单汇总成本'] = lingxing_order_data['平台单号'].map(order_total_cost)

#按系统单号分组求和 单个订单成本 列,用于后续广告费按订单成本比例进行分摊
lingxing_order_data['分摊广告费金额'] = (lingxing_order_data['单个订单成本'] /
                                      lingxing_order_data['订单汇总成本'] *
                                      lingxing_order_data['广告费汇总'])

ads_fee.to_excel(os.path.join(result_path,"广告费-需额外整理.xlsx"), engine='xlsxwriter', index=False)
lingxing_order_data.to_excel(os.path.join(result_path,"广告费-领星订单对应明细.xlsx"), engine='xlsxwriter', index=False)

#筛选费用标签类型：“需额外整理”且核算科目为“海外仓运费”，复制到一个新表，表格命名为海外仓运费-需额外整理
overseas_ship_fee = final_merged_data[(final_merged_data['费用标签类型'] == '需额外整理') & (final_merged_data['核算科目'] == '海外仓运费')]

#复制海外仓运费-需额外整理中的订单编号去领星-订单管理-批量搜索订单编号--不合并单元格导出-导出字段全选（不勾图片）--导出，导出来数据放到领星对应订单明细里面

lingxing_order_data=lingxing_order_data_initial.copy()

# 获取所有在 overseas_ship_fee 中存在的订单编号（即海外仓运费订单）
overseas_ship_order_numbers = overseas_ship_fee['订单编号'].unique()

# 在 lingxing_order_data 中筛选出平台单号存在于 overseas_ship_order_numbers 的行
lingxing_order_data = lingxing_order_data[lingxing_order_data['平台单号'].isin(overseas_ship_order_numbers)]
print(len(overseas_ship_order_numbers),len(lingxing_order_data))

#确保系统单号和商品ID是字符串类型，免得被科学计数法了
lingxing_order_data['系统单号'] = lingxing_order_data['系统单号'].fillna('').astype(str)
lingxing_order_data['订单商品ID'] = lingxing_order_data['订单商品ID'].fillna('').astype(str)
lingxing_order_data['ASIN/商品Id'] = lingxing_order_data['ASIN/商品Id'].fillna('').astype(str)
lingxing_order_data['参考号'] = lingxing_order_data['参考号'].fillna('').astype(str)

# 删除 SKU 为 NaN 或 空字符串 的行
lingxing_order_data = lingxing_order_data[lingxing_order_data['SKU'].notna() & (lingxing_order_data['SKU'] != '')]

#创建一个平台单号到海外仓运费的映射字典
ad_fee_sum_dict = overseas_ship_fee.groupby('订单编号')['交易总金额（RMB）'].sum().to_dict()
lingxing_order_data['海外仓运费汇总'] = lingxing_order_data['平台单号'].map(ad_fee_sum_dict)

#读取产品信息表（用于映射产品采购单价）
product_file_path = os.path.join(auxiliary_data_path,r"领星+店小秘产品信息表.xlsx")
product_information = pd.read_excel(product_file_path, sheet_name='普通+组合产品',header=1)

#获取 SKU 对应的「单个产品成本价格」，用于计算「单个产品成本价格」,构建 SKU 到「采购单价（核算）」的映射字典
sku_to_cost_price = product_information.set_index('*SKU')['采购单价（核算）'].to_dict()
lingxing_order_data['单个产品成本价格'] = lingxing_order_data['SKU'].map(sku_to_cost_price)

# 使用 map 添加「单个产品成本价格」列,为每个订单商品匹配其对应的产品成本价，用于后续成本核算
lingxing_order_data['单个订单成本'] = lingxing_order_data['数量'] * lingxing_order_data['单个产品成本价格']

#单个订单成本 = 数量 * 单个产品成本价格,用于统计每个商品在订单中的总成本
order_total_cost = lingxing_order_data.groupby('平台单号')['单个订单成本'].sum().to_dict()
lingxing_order_data['订单汇总成本'] = lingxing_order_data['平台单号'].map(order_total_cost)

#按系统单号分组求和 单个订单成本 列,用于后续海外仓运费按订单成本比例进行分摊
lingxing_order_data['分摊海外仓运费金额'] = (lingxing_order_data['单个订单成本'] /
                                      lingxing_order_data['订单汇总成本'] *
                                      lingxing_order_data['海外仓运费汇总'])

overseas_ship_fee.to_excel(os.path.join(result_path,"海外仓运费-需额外整理.xlsx"), engine='xlsxwriter', index=False)
lingxing_order_data.to_excel(os.path.join(result_path,"海外仓运费-领星订单对应明细.xlsx"), engine='xlsxwriter', index=False)


# 读取ebay返款表.xlsx
ebay_cashback=pd.read_excel(os.path.join(auxiliary_data_path,"ebay海外退件+返款表.xlsx"),sheet_name='返款表')
# 筛选当前核算月的数据
ebay_cashback = ebay_cashback[ebay_cashback['返款时间'].dt.to_period('M') == f'{prev_year}-{prev_month:02d}']
# 币种匹配当月汇率，汇率来自最新负责人-确认表
ebay_cashback["汇率"] = ebay_cashback["币种"].map(currency_to_rate)
# 金额（RMB) 等于 净额（外币） * 汇率
ebay_cashback["金额（RMB）"] = ebay_cashback["净额（外币）"] * ebay_cashback["汇率"]
ebay_cashback["核算月份"]=f"{prev_year}-{prev_month}"
#生成excel，后续传数跨境
ebay_cashback.to_excel(os.path.join(result_path,"ebay返款表.xlsx"), engine='xlsxwriter', index=False)
#%%
# 读取ebay海外仓退件表.xlsx,订单号设置为字符串类型
ebay_overseas_return=pd.read_excel(os.path.join(auxiliary_data_path,"ebay海外退件+返款表.xlsx"),dtype={'订单号': str},sheet_name='海外退件表')

## 登记日期转为日期类型
ebay_overseas_return['登记日期'] = pd.to_datetime(ebay_overseas_return['登记日期'], errors='coerce')
# 筛选出登记日期属于上个月的数据
ebay_overseas_return = ebay_overseas_return[
    (ebay_overseas_return['登记日期'].dt.year == prev_year) &
    (ebay_overseas_return['登记日期'].dt.month == prev_month)
]
# 去除“订单号”为空的行，
ebay_overseas_return = ebay_overseas_return[ebay_overseas_return['订单号'].notna() & (ebay_overseas_return['订单号'] != '')]
# 确保数量 和单个成本是数值
ebay_overseas_return['数量'] = ebay_overseas_return['数量'].astype(float)
ebay_overseas_return['单个成本'] = ebay_overseas_return['单个成本'].astype(float)

lingxing_order_data=lingxing_order_data_initial.copy()

#设置系统单号为字符串类型，防止发货时间匹配不上
lingxing_order_data['系统单号'] = lingxing_order_data['系统单号'].astype(str)

#新增一列发货时间,按订单号匹配出发货时间
order_id_to_ship_time = lingxing_order_data.drop_duplicates(subset=['系统单号'], keep='first').set_index('系统单号')['发货时间'].to_dict()
ebay_overseas_return["发货时间"] = ebay_overseas_return["订单号"].map(order_id_to_ship_time)

#直接下载所有的售后工单，一单单查烦死，新增一列退款时间,按订单号匹配出退款时间
lingxing_after_sale_order_data=pd.read_excel(os.path.join(auxiliary_data_path,"领星导出-售后工单.xlsx"))

# 筛选售后类型为"仅退款"、"退货退款"
lingxing_after_sale_order_data = lingxing_after_sale_order_data[
    lingxing_after_sale_order_data['售后类型'].isin(['仅退款', '退货退款'])
]

#系统单号/平台单号的格式为"103579668858496137\n01-13181-81658" ，需要进一步拆分
lingxing_after_sale_order_data[["系统单号", "平台单号"]] = lingxing_after_sale_order_data["系统单号/平台单号"].str.split("\n", expand=True)
#创建人/创建时间的格式为"龙莹\n2025-07-05 09:59:29" ，需要进一步拆分
lingxing_after_sale_order_data[["创建人", "创建时间"]] = lingxing_after_sale_order_data["创建人/创建时间"].str.split("\n", expand=True)

#新增一列退款时间,按订单号匹配出退款时间
order_id_to_refund_time = lingxing_after_sale_order_data.drop_duplicates(subset=['系统单号'], keep='first').set_index('系统单号')['创建时间'].to_dict()
ebay_overseas_return["退款时间"] = ebay_overseas_return["订单号"].map(order_id_to_refund_time).fillna('')

#获取 SKU 对应的「状态」，用于新增一列「状态」,构建 SKU 到「状态」的映射字典
sku_to_product_status = product_information.set_index('*SKU')['状态'].to_dict()
ebay_overseas_return["状态"]=ebay_overseas_return['SKU'].map(sku_to_product_status)

#增加一列补贴金额，数量*单个成本
# 对“退回上架”、“重新上架”按数量 × 单个成本 × 100% 计算补贴金额
mask_1 = ebay_overseas_return['处理方法'].isin(['退回上架', '重新上架'])
ebay_overseas_return.loc[mask_1, '补贴金额'] = ebay_overseas_return.loc[mask_1, '数量'] * ebay_overseas_return.loc[mask_1, '单个成本'] * 1.0

# 对“退回退款”按数量 × 单个成本 × 60% 计算补贴金额
mask_2 = ebay_overseas_return['处理方法'] == '退回退款'
ebay_overseas_return.loc[mask_2, '补贴金额'] = ebay_overseas_return.loc[mask_2, '数量'] * ebay_overseas_return.loc[mask_2, '单个成本'] * 0.6

#加一个“核算月份”。发货时间、退款时间都齐的话，计入当月，如果无退款时间，留空白
# 确保发货时间和退款时间是 datetime 类型
ebay_overseas_return['发货时间'] = pd.to_datetime(ebay_overseas_return['发货时间'], errors='coerce')
ebay_overseas_return['退款时间'] = pd.to_datetime(ebay_overseas_return['退款时间'], errors='coerce')

# 新增「核算月份」列，默认为空
ebay_overseas_return['核算月份'] = ''

# 只有当发货时间和退款时间都不为空时，才填写核算月份
valid_mask = (ebay_overseas_return['发货时间'].notna() & ebay_overseas_return['退款时间'].notna() )
ebay_overseas_return.loc[valid_mask, '核算月份'] = f"{prev_year}-{prev_month}"
# #生成excel，后续传数跨境
ebay_overseas_return.to_excel(os.path.join(result_path,"ebay海外仓退件表.xlsx"), engine='xlsxwriter', index=False)

#  读取谷仓重新上架退件表
gucong_relisted_returns=pd.read_excel(os.path.join(auxiliary_data_path,"谷仓重新上架退件表.xlsx"))
print(gucong_relisted_returns.columns)
# 筛选出ebay平台数据
gucong_relisted_returns = gucong_relisted_returns[gucong_relisted_returns["平台"].str.contains("eBay", case=False, na=False)]

#原订单参考号匹配系统单号获取创建时间
gucong_relisted_returns["发货时间"] = gucong_relisted_returns["原订单参考号"].map(order_id_to_ship_time)
# 复制当月原订单参考号，“退款时间”，领星-客服-售后工单--平台选择ebay--创建时间选择近3个月--下载原订单参考号匹配系统单号获取创建时间
# lingxing_order_data=lingxing_order_data_initial.copy()
#“退款时间”，领星-客服-售后工单--平台选择ebay--创建时间选择近3个月--下载
gucong_relisted_returns["退款时间"] = gucong_relisted_returns["原订单参考号"].map(order_id_to_refund_time).fillna('')

#“状态”：领星SKU对应“领星-店小秘产品信息表”SKU拉取状态
gucong_relisted_returns["状态"] = gucong_relisted_returns["领星SKU"].map(sku_to_product_status)

#“单个成本”:领星SKU对应“领星-店小秘产品信息表”SKU拉取采购单价（核算）
gucong_relisted_returns["单个成本"] = gucong_relisted_returns["领星SKU"].map(sku_to_cost_price)

#“补贴金额”:上架数量*单个成本
gucong_relisted_returns["补贴金额"] = gucong_relisted_returns["上架数量"] * gucong_relisted_returns["单个成本"]

#发货时间、退款时间不为空的计入月份填当月，退款时间空白计入月份计入月份为空白
# 确保发货时间和退款时间是 datetime 类型
gucong_relisted_returns['发货时间'] = pd.to_datetime(gucong_relisted_returns['发货时间'], errors='coerce')
gucong_relisted_returns['退款时间'] = pd.to_datetime(gucong_relisted_returns['退款时间'], errors='coerce')

# 新增「核算月份」列，默认为空
gucong_relisted_returns['计入月份'] = ''

# 只有当发货时间和退款时间都不为空时，才填写核算月份
valid_mask = (gucong_relisted_returns['发货时间'].notna() & gucong_relisted_returns['退款时间'].notna() )
gucong_relisted_returns.loc[valid_mask, '计入月份'] = f"{prev_year}-{prev_month}"

# #生成excel，后续传数跨境
gucong_relisted_returns.to_excel(os.path.join(result_path,"谷仓退件补贴.xlsx"), engine='xlsxwriter', index=False)

# 读取订单管理
lingxing_order_data=lingxing_order_data_initial.copy()
lingxing_order_data['发货时间'] = pd.to_datetime(lingxing_order_data['发货时间'])

# 筛选上个月的数据
lingxing_order_data = lingxing_order_data[
    (lingxing_order_data['发货时间'].dt.year == prev_year) &
    (lingxing_order_data['发货时间'].dt.month == prev_month)
]
lingxing_order_data['平台单号'] = lingxing_order_data['平台单号'].astype(str)
lingxing_order_data['系统单号'] = lingxing_order_data['系统单号'].astype(str)
lingxing_order_data['运单号'] = lingxing_order_data['运单号'].astype(str)

# 新增“订单类型2”，填写：付款订单
lingxing_order_data["订单类型2"]="付款订单"

#按照“店铺”和“系统单号”分组统计数量。
lingxing_order_data['系统单号个数'] = lingxing_order_data.groupby(["店铺", "系统单号"]).transform('size')
#按照“店铺”和“平台单号”分组统计数量。
lingxing_order_data['平台单号个数'] = lingxing_order_data.groupby(["店铺", "系统单号"]).transform('size')

#汇率：按“订单币种”去拉取最新负责人-确认版中汇率“币种二字码”对应的最新月份汇率
lingxing_order_data["汇率"] = lingxing_order_data["订单币种"].map(currency_to_rate)

#单个产品成本价格：按“SKU”去拉取领星+店小秘产品信息表中--普通+组合产品附表中的“*SKU”对应的“采购单价（核算）”
lingxing_order_data['单个产品成本价格'] = lingxing_order_data['SKU'].map(sku_to_cost_price).fillna(0)

#获取 SKU 对应的「单个产品重量」，构建 SKU 到「单品毛重」的映射字典,后续计算「单个订单重量」
sku_to_product_weight = product_information.set_index('*SKU')['单品毛重'].to_dict()

#单个产品重量 （匹配产品表）：按“SKU”去拉取领星+店小秘产品信息表中--普通+组合产品附表中的“*SKU”对应的“单品毛重”
lingxing_order_data['单个产品重量（匹配产品表）'] = lingxing_order_data['SKU'].map(sku_to_product_weight).fillna(0)

#单个订单重量：公式正常往下拉，数量*单个产品重量（匹配产品表）
lingxing_order_data['单个订单重量']=lingxing_order_data['数量']*lingxing_order_data['单个产品重量（匹配产品表）']

#重量汇总：公式正常往下拉，根据“系统单号”“店铺”“订单类型2” 汇总“单个订单重量”
lingxing_order_data['重量汇总'] = lingxing_order_data.groupby(['系统单号', '店铺', '订单类型2'])['单个订单重量'].transform('sum')

#销售收入：（商品金额+商品客付运费）*汇率
lingxing_order_data['销售收入'] = (lingxing_order_data['商品金额'] + lingxing_order_data['商品客付运费']) * lingxing_order_data['汇率']

#交易佣金：-商品交易费*汇率
lingxing_order_data['交易佣金'] = -lingxing_order_data['商品交易费'] * lingxing_order_data['汇率']

#税费收入：商品客付税费*汇率
lingxing_order_data['税费收入'] = lingxing_order_data['商品客付税费'] * lingxing_order_data['汇率']

#扣除客户支付税费：-商品客付税费*汇率
lingxing_order_data['扣除客户支付税费'] = -lingxing_order_data['商品客付税费'] * lingxing_order_data['汇率']

#读取谷线下付款申请表-售后订单，修改部分售后备品仓的采购成本，其余采购成本为0
offline_payment_after_sales_orders=pd.read_excel(os.path.join(auxiliary_data_path,"线下付款申请表-售后订单.xlsx"),sheet_name="线下付款申请表")

#初始化采购成本的值
lingxing_order_data['采购成本'] = -lingxing_order_data["商品出库成本"]

#筛选“发货仓库”为“售后备品仓”，设置采购成本为0
lingxing_order_data.loc[lingxing_order_data["发货仓库"] == "售后备品仓", "采购成本"] = 0

#线下付款申请表，根据系统订单号匹配对应的采购成本，涉及SKU：Zhanglijun，Luoqian，Wenyanxia，Yuliu，Qinyuan，Huangbinhua
valid_payments = offline_payment_after_sales_orders.dropna(subset=['系统订单号（售后订单需要填写）'])
order_id_to_payment_amount = (valid_payments.groupby('系统订单号（售后订单需要填写）')['付款金额'].sum().to_dict())

# 创建一个 mask，标识哪些行的“系统订单号”存在于售后订单表中
matched_rows = lingxing_order_data['系统单号'].isin(order_id_to_payment_amount.keys())

# 只对这些匹配的行更新“采购成本”
lingxing_order_data.loc[matched_rows, '采购成本'] = lingxing_order_data.loc[matched_rows, '系统单号'].map(order_id_to_payment_amount)

#筛选“发货仓库”美国售后**仓/美国移除**：负数“数量”*“单个产品成本价格”*60%，且“正常单/售后单区分”填写售后单
after_sale_condition = lingxing_order_data['发货仓库'].str.contains(r'美国售后.*仓|美国移除.*', regex=True)

# 对符合after_sale_condition条件的行设置「采购成本」
lingxing_order_data.loc[after_sale_condition, '采购成本'] = (-lingxing_order_data.loc[after_sale_condition, '数量'] * lingxing_order_data.loc[after_sale_condition, '单个产品成本价格']) * 0.6

# 对符合after_sale_condition条件的行设置「正常单/售后单区分」为售后单
lingxing_order_data.loc[after_sale_condition, '正常单/售后单区分'] = '售后单'




# 筛选“发货仓库”是 Amazon*仓 的订单
amazon_warehouse_condition = lingxing_order_data['发货仓库'].str.contains(r'Amazon.*仓',na=False,regex=True)
# 获取符合条件的系统单号
amazon_order_ids = lingxing_order_data.loc[amazon_warehouse_condition, '运单号']
lingxing_search_order_ids = amazon_order_ids.unique()

print("共", len(lingxing_search_order_ids), "条 Amazon 仓订单号：")
for oid in lingxing_search_order_ids:
    print(oid)
print("----------------------------------------------------------------")
# 每199个订单号分组
batch_size = 199
order_id_batches = [lingxing_search_order_ids[i:i + batch_size] for i in range(0, len(lingxing_search_order_ids), batch_size)]
# 打印每个批次的数量和内容
file_order_dict_list = []
for idx, batch in enumerate(order_id_batches):
    file_name = os.path.join(auxiliary_data_path,f"多渠道订单列表-发货仓库为Amazon仓第{idx+1}批.xlsx")
    file_order_dict = {
        file_name: list(batch)
    }
    file_order_dict_list.append(file_order_dict)

for i in file_order_dict_list:
    download_file_path=list(i.keys())[0]
    order_id_lsit=list(i.values())[0]
    # print("="*40)
    # print(download_file_path)
    # print('\n'.join(order_id_lsit))
# 逐个打印需要下载的订单数据，199个/批次
# print('\n'.join(list(file_order_dict_list[1].values())[0]))
#所有领星下载的多渠道订单列表-发货仓库为Amazon仓第1批.xlsx
file_paths = [list(d.keys())[0] for d in file_order_dict_list]
#合并下载的领星多渠道订单列表
multi_channel_orders = pd.concat(
    [pd.read_excel(file_path) for file_path in file_paths],
    ignore_index=True
)
multi_channel_orders=multi_channel_orders[['卖家订单号', 'ASIN', 'MSKU', '亚马逊订单号']]


multi_channel_orders['卖家订单号'] = multi_channel_orders['卖家订单号'].astype(str)
multi_channel_orders['亚马逊订单号'] = multi_channel_orders['亚马逊订单号'].astype(str)

multi_channel_orders.rename(columns={
    '卖家订单号': '运单号',
    'ASIN': '亚马逊ASIN',
    'MSKU': '亚马逊MSKU',
    '亚马逊订单号': '亚马逊订单号'
}, inplace=True)
print("前",len(lingxing_order_data))
lingxing_order_data = lingxing_order_data.merge(
    multi_channel_orders,
    on='运单号',
    how='left'  # 保留原数据，即使没有匹配项
)
print("后",len(lingxing_order_data))
# 复制亚马逊订单号到领星--亚马逊平台--财务--成本计价--业务编码（粘贴亚马逊订单号）搜索--导出
# 每1999个订单号分组
lingxing_order_data['亚马逊订单号'] = lingxing_order_data['亚马逊订单号'].fillna('').astype(str)
lingxing_search_order_ids = lingxing_order_data["亚马逊订单号"].unique()

# 打印唯一订单号的数量
print(f"唯一订单号数量: {len(lingxing_search_order_ids)}")

# 打印所有唯一订单号（每行一个）
print("具体订单号列表:")
for order_id in lingxing_search_order_ids:
    print(order_id)



