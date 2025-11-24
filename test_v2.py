import dataStandard_v2
import weight_v2
import threshold_v2
import pandas as pd
import config

# 加载Excel文件
df = pd.read_excel("./data/final_result.xlsx", engine="openpyxl")

df["更新时间"] = df["date"]
df["不合格率"] = 1 - df["合格率"]

need_df_cols = ["工位", "更新时间", "检验成本", "不合格率", "返工成本", "报废成本"]
df = df[need_df_cols].copy()

# 获取标准化指标值（包含时间维度）
try:
    standardData = dataStandard_v2.Normaliz(df, config.beta, config.log_c, config.log_d)
    # 从standardData获取更新时间列表
    update_times = standardData["更新时间"].unique()
    update_times = sorted(update_times)  # 确保时间有序
    print(f"找到 {len(update_times)} 个更新时间点")

except Exception as e:
    print(f"获取标准化数据失败：{str(e)}")

# 获取最终权重（包含时间维度）
try:
    final_weights = weight_v2.CombinedWeight(df, config.expert_weights, config.expert_weights_percent)
    # 检查权重数据的时间点是否与标准化数据匹配
    weight_times = set(final_weights["更新时间"].unique())
    standard_times = set(update_times)
    if not standard_times.issubset(weight_times):
        missing_times = standard_times - weight_times
        print(f"警告：权重数据缺少以下时间点的数据: {missing_times}")

except Exception as e:
    print(f"获取权重失败：{str(e)}")

import pandas as pd

# --------------------------
# 步骤1：权重数据列重命名（避免与指标列冲突）
# --------------------------
# 为权重的4个指标列添加“_权重”后缀，明确区分“指标值”和“权重”
final_weights_renamed = final_weights.rename(
    columns={
        "检验成本": "检验成本_权重",
        "不合格率": "不合格率_权重",
        "返工成本": "返工成本_权重",
        "报废成本": "报废成本_权重"
    }
)

# --------------------------
# 步骤2：按“更新时间”合并两个数据集
# --------------------------
# 左连接：以标准化数据的时间为准，确保所有工位的指标都能匹配到对应时间的权重
# （若某时间无权重，权重列会显示NaN，可通过后续代码检查）
merged_data = pd.merge(
    left=standardData,          # 标准化指标数据（含工位、时间、指标值）
    right=final_weights_renamed, # 重命名后的权重数据（含时间、权重）
    on="更新时间",               # 合并键：更新时间
    how="left"                  # 左连接：保留所有标准化数据行
)

# --------------------------
# 步骤3：检查是否存在权重缺失（可选但推荐）
# --------------------------
missing_weight_rows = merged_data[merged_data["检验成本_权重"].isna()]
if not missing_weight_rows.empty:
    missing_times = missing_weight_rows["更新时间"].unique()
    print(f"警告:以下时间无对应权重,加权值会显示NaN:{missing_times}")
else:
    print("所有时间的权重均匹配成功，可继续计算加权值")

# --------------------------
# 步骤4：计算每个指标的“标准化值 × 权重”（加权值）
# --------------------------
merged_data["检验成本_加权值"] = merged_data["检验成本"] * merged_data["检验成本_权重"]
merged_data["不合格率_加权值"] = merged_data["不合格率"] * merged_data["不合格率_权重"]
merged_data["返工成本_加权值"] = merged_data["返工成本"] * merged_data["返工成本_权重"]
merged_data["报废成本_加权值"] = merged_data["报废成本"] * merged_data["报废成本_权重"]
merged_data["结果"] = 100-(merged_data["检验成本_加权值"] + merged_data["不合格率_加权值"] + merged_data["返工成本_加权值"] + merged_data["报废成本_加权值"])*100

# --------------------------
# 步骤5：整理最终结果（可选，按时间+工位排序）
# --------------------------
final_result = merged_data.sort_values(by=["更新时间", "工位"]).reset_index(drop=True)

# --------------------------
# 步骤6：根据阈值，打分
# --------------------------
score, _ = threshold_v2.GradeThreshold(final_result, config.CV_High, config.CV_Low, config.excel_save_path)