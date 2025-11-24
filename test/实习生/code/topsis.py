import pandas as pd
import numpy as np
 
# 读取Excel文件
file_path = 'data.xlsx'
df = pd.read_excel(file_path, engine='openpyxl')
df_selected = df.iloc[:, 1:5]
 
# 数据标准化
def normalize_data(df):
    norm_df = df.copy()
    for column in df.columns:
        min_val = df[column].min()
        max_val = df[column].max()
        norm_df[column] = (df[column] - min_val) / (max_val - min_val)
    return norm_df
 
norm_df = normalize_data(df_selected)
 
# 计算信息熵
def calculate_entropy(norm_df):
    entropy = np.zeros(norm_df.shape[1])
    for i, column in enumerate(norm_df.columns):
        p = norm_df[column] / norm_df[column].sum()
        entropy[i] = -np.sum(p * np.log(p + 1e-9)) / np.log(len(norm_df))
    return entropy
 
entropy = calculate_entropy(norm_df)
 
# 计算差异系数
d = 1 - entropy
 
# 计算权重
weights = d / d.sum()
 
# 输出权重
weight_dict = dict(zip(df_selected.columns, weights))
print("各指标的权重如下：")
for indicator, weight in weight_dict.items():
    print(f"{indicator}: {weight:.4f}")