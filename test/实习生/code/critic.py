import pandas as pd
import numpy as np
 
# 读取Excel文件中的数据
df = pd.read_excel('test_data.xlsx', engine='openpyxl')
 

# 数据无量纲化处理（从第二列开始处理四个指标）
def normalize_data(df):
    normalized_df = pd.DataFrame()
    selected_columns = df.columns[1:5]  # 从第二列开始，处理四列
    for i, column in enumerate(selected_columns):
        min_val = df[column].min()
        max_val = df[column].max()
        #if i < 3:  # 前三个指标做正向化处理
        #normalized_df[column] = (df[column] - min_val) / (max_val - min_val)
        #else:  # 第四个指标做逆向化处理
        normalized_df[column] = (max_val - df[column]) / (max_val - min_val)
    return normalized_df
 
normalized_df = normalize_data(df)
 
# 计算指标的变异性（标准差）
std_devs = normalized_df.std()
 
# 计算指标的冲突性（相关系数）
correlation_matrix = normalized_df.corr()
conflict_values = correlation_matrix.apply(lambda x: 1 - x).sum()
 
# 计算信息量
information_values = std_devs * conflict_values
 
# 计算权重
weights = information_values / information_values.sum()
weights_array = weights.values  # 输出是形状为 (4,) 的 ndarray


# 输出权重
print("指标权重如下：")
print(weights_array)

# 定义因素集和评语集
factors = ['返工总用时', '检验用时', '单价', '总价']
comments = ['优', '中', '差']


# 定义隶属函数

# 返工总用时隶属度
def fan_time_best(x):
    if x < 10:
        return 1
    if 10 <= x <= 27:
        return (27-x) / 17
    if x > 27:
        return 0

def fan_time_medium(x):
    if x < 27:
        return 0
    if 27 <= x <= 56:
        return (x-27) / 29
    if 56 < x < 62:
        return (62-x) / 6
    if x >= 62:
        return 0

def fan_time_bad(x):
    if x <= 62:
        return 0
    if x > 62:
        return (720-x) / 658


# 检验用时
def jian_time_best(x):
    if x < 5:
        return 1
    if 5 <= x <= 8:
        return (8-x) / 3
    if x > 8:
        return 0

def jian_time_medium(x):
    if x < 8:
        return 0
    if 8 <= x <= 10:
        return (x-8) / 2
    if 10 < x < 13:
        return (13-x) / 3
    if x >= 13:
        return 0

def jian_time_bad(x):
    if x <= 13:
        return 0
    if x > 13:
        return (30-x) / 17


# 单价
def d_price_best(x):
    if x < 250:
        return 1
    if 250 <= x <= 300:
        return (300-x) / 50
    if x > 300:
        return 0

def d_price_medium(x):
    if x < 300:
        return 0
    if 300 <= x <= 400:
        return (x-300) / 100
    if 400 < x < 500:
        return (500-x) / 3
    if x >= 500:
        return 0

def d_price_bad(x):
    if x <= 500:
        return 0
    if x > 500:
        return 1


# 总价
def z_price_best(x):
    if x < 20:
        return 1
    if 20 <= x <= 57:
        return (57-x) / 37
    if x > 57:
        return 0

def z_price_medium(x):
    if x < 57:
        return 0
    if 57 <= x <= 107:
        return (x-57) / 50
    if 107 < x < 141:
        return (141-x) / 343
    if x >= 141:
        return 0

def z_price_bad(x):
    if x <= 141:
        return 0
    if x > 141:
        return (2509-x) / 2368

# 构建模糊隶属度矩阵函数（每人对应一个4x3矩阵）
def compute_membership_matrix(person):
    mtx = np.zeros((4, 3))  # 4 个指标 - 3 个等级
    mtx[0] = [fan_time_best(person[0]), fan_time_medium(person[0]), fan_time_bad(person[0])]
    mtx[1] = [jian_time_best(person[1]), jian_time_medium(person[1]), jian_time_bad(person[1])]
    mtx[2] = [d_price_best(person[2]), d_price_medium(person[2]), d_price_bad(person[2])]
    mtx[3] = [z_price_best(person[3]), z_price_medium(person[3]), z_price_bad(person[3])]
    return mtx


# 读取数据
df = pd.read_excel('test_data.xlsx', engine='openpyxl')  # 假设第2~5列为四个指标
data = df.iloc[:, 1:5].values  # 每行是一个人的四项指标值
# 为所有人计算模糊矩阵
all_membership_matrices = [compute_membership_matrix(person) for person in data]

# 打印第一个人的模糊矩阵
print("第一项的模糊隶属矩阵：")
print(all_membership_matrices[3])  # 形状为 (4,3)

# 计算加权后的每个因素的综合权重  点乘方式dot 可修改
factor_weights = np.dot(weights, all_membership_matrices[3])  # 得到形如 (4,) 的向量
print("第一项的评分：")
print(factor_weights) 

# 找出最大值的索引
max_index = np.argmax(factor_weights)

# 映射索引到评语
comments = ['优', '中', '差']
final_comment = comments[max_index]

print("最终等级：", final_comment)

# 添加到原表格中并保存
#df['综合评语'] = final_comments
#f.to_excel('data_with_comments.xlsx', index=False)

# 用于存储每人的等级结果
final_comments = []

for person in data:
    # 获取每个人的模糊矩阵
    membership = compute_membership_matrix(person)  # shape: (4, 3)
    
    # 对每一行（每个因素）使用权重加权，得到形如 (4,) 的得分向量
    factor_scores = np.dot(weights, membership)  # 权重 ⋅ 每列隶属度
    
    # 找出得分中最大的维度所对应的等级
    max_index = np.argmax(factor_scores)
    comment = comments[max_index]
    
    # 存储结果
    final_comments.append(comment)

# 添加结果到原表格
df['最终等级'] = final_comments

# 保存新表格
df.to_excel('data_with_final_grades.xlsx', index=False)
print("综合等级结果已保存")