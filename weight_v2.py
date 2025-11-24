import pandas as pd
import warnings
import config

# 定义函数：结合专家打分和 CRITIC 权重
def CombinedWeight(df, expert_weights, expert_weights_percent):
    # 按更新时间排序
    df = df.sort_values(by="更新时间")
    z_values = sorted(df["更新时间"].unique())  # 时间从小到大
    
    # 计算专家权重（所有时间点保持一致）
    expert_df = pd.DataFrame(expert_weights)
    expert_mean_weights = expert_df.mean()  # 专家权重平均值，固定不变
    
    # 存储每个时间点的合并权重
    combined_weights_list = []
    
    # 直接遍历每个时间点
    for current_time in z_values:
        # 获取截至当前时间点的所有历史数据（包括当前时间点）
        current_data = df[df['更新时间'] <= current_time].copy()
        
        # 1. 极差标准化，处理除零情况
        cols = ['检验成本', '不合格率', '返工成本', '报废成本']
        max_vals = current_data[cols].max()
        min_vals = current_data[cols].min()
        ranges = max_vals - min_vals
        # 当极差为0时，避免除零错误，标准化为0
        normalized_df = (current_data[cols] - min_vals) / ranges.where(ranges != 0, 1)
        
        # 2. 计算标准差
        std_dev = normalized_df.std()
        # 处理只有一个数据点时标准差为NaN的情况
        std_dev = std_dev.fillna(0)
        
        # 3. 计算冲突性矩阵及总和
        if len(current_data) < 2:
            # 数据不足2条时无法计算相关性，默认冲突为0
            conflict_sum = pd.Series([0]*len(cols), index=cols)
        else:
            correlation_matrix = normalized_df.corr()
            conflict_matrix = 1 - correlation_matrix  # 冲突性矩阵
            conflict_sum = conflict_matrix.sum()
        
        # 4. 计算CRITIC权重
        critic_score = std_dev * conflict_sum
        # 处理分数为0的特殊情况（平均分配权重）
        if critic_score.sum() == 0:
            critic_weights = pd.Series([1/len(critic_score)]*len(critic_score), 
                                      index=critic_score.index)
        else:
            critic_weights = critic_score / critic_score.sum()
        
        # 5. 合并权重（专家权重与CRITIC权重各占50%）
        combined = expert_weights_percent * expert_mean_weights + (1 - expert_weights_percent) * critic_weights
        combined = combined / combined.sum()  # 确保权重和为1
        
        # 存储当前时间点的权重，包含对应的时间
        weight_dict = {"更新时间": current_time}
        weight_dict.update(combined.round(4).to_dict())
        combined_weights_list.append(weight_dict)
    
    # 将列表转换为DataFrame
    weight_df = pd.DataFrame(combined_weights_list)
    return weight_df

def main():
    # Load the Excel file
    df = pd.read_excel("./data/final_result.xlsx", engine="openpyxl")
    
    df["更新时间"] = df["date"]
    df["不合格率"] = 1 - df["合格率"]

    need_df_cols = ["工位", "更新时间", "检验成本", "不合格率", "返工成本", "报废成本"]
    df = df[need_df_cols].copy()

    # 专家打分权重：每个专家打一次分
    expert_weights = [
        {'检验成本': 0.25, '不合格率': 0.26, '返工成本': 0.22, '报废成本': 0.27}, # 专家1
        {'检验成本': 0.24, '不合格率': 0.26, '返工成本': 0.24, '报废成本': 0.26}, # 专家2
        {'检验成本': 0.23, '不合格率': 0.27, '返工成本': 0.25, '报废成本': 0.25}, # 专家3
        {'检验成本': 0.22, '不合格率': 0.28, '返工成本': 0.25, '报废成本': 0.25}, # 专家4
        {'检验成本': 0.21, '不合格率': 0.29, '返工成本': 0.25, '报废成本': 0.25}, # 专家5
        {'检验成本': 0.20, '不合格率': 0.30, '返工成本': 0.25, '报废成本': 0.25}, # 专家6
        {'检验成本': 0.19, '不合格率': 0.31, '返工成本': 0.25, '报废成本': 0.25}, # 专家7
        {'检验成本': 0.18, '不合格率': 0.32, '返工成本': 0.25, '报废成本': 0.25}, # 专家8
        {'检验成本': 0.17, '不合格率': 0.33, '返工成本': 0.25, '报废成本': 0.25}, # 专家9
        {'检验成本': 0.20, '不合格率': 0.30, '返工成本': 0.25, '报废成本': 0.25}, # 专家10
    ]

    expert_weights_percent = 0.5
    final_weights = CombinedWeight(df, expert_weights, expert_weights_percent)

    # 设置显示所有行（取消行数限制）
    pd.set_option('display.max_rows', None)
    # 设置显示所有列（取消列数限制）
    pd.set_option('display.max_columns', None)
    # 设置列宽（避免内容被截断）
    pd.set_option('display.max_colwidth', None)
    pd.set_option('display.float_format', '{:.6f}'.format)

    # 打印完整DataFrame
    print(final_weights)

    # for idx, row in final_weights.iterrows():
    #     print(f"\n时间点 {idx + 1}: {row['更新时间']:.0f}")
    #     print("-" * 30)
    #     for indicator in ["检验成本", "合格率", "返工成本", "报废成本"]:
    #         print(f"{indicator}: {row[indicator]:.4f}")

if __name__ == "__main__":
    main()