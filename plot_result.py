import os
import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# ================== 可配置区 ==================
file_path = "./result/test_result.xlsx"   # 输入Excel 文件路径
output_dir = "./charts_result"                  # 图片输出目录
os.makedirs(output_dir, exist_ok=True)
save_recalc_excel = True                 # 是否导出带映射结果与新等级的Excel
recalc_excel_path = "./result/test_result_mapped.xlsx"
# ============================================

# ---- 读表 ----
df = pd.read_excel(file_path, engine="openpyxl")

# ---- 列名 ----
time_col = "更新时间" 
station_col = "工位"
result_col = "结果"
t_high_col = "T_high"
t_low_col  = "T_low"

# ---- 类型转换 ----
df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
df[result_col] = pd.to_numeric(df[result_col], errors="coerce")
df[t_high_col] = pd.to_numeric(df[t_high_col], errors="coerce")
df[t_low_col]  = pd.to_numeric(df[t_low_col], errors="coerce")

# ---- 统一阈值：中位数 ----
T_HIGH = df[t_high_col].median(skipna=True)
T_LOW  = df[t_low_col].median(skipna=True)

if T_LOW >= T_HIGH:
    raise ValueError(f"阈值不合理: T_LOW({T_LOW}) 应小于 T_HIGH({T_HIGH})。")

df["T_high_unified"] = float(T_HIGH)
df["T_low_unified"]  = float(T_LOW)

# 基于旧阈值的相对位置，映射到统一阈值下的新“结果”
# r_new = T_LOW + (结果 - T_low_old) * ((T_HIGH - T_LOW) / (T_high_old - T_low_old))
vals     = pd.to_numeric(df[result_col], errors="coerce")
low_old  = pd.to_numeric(df[t_low_col],  errors="coerce")
high_old = pd.to_numeric(df[t_high_col], errors="coerce")

new_range = float(T_HIGH - T_LOW)
old_range = (high_old - low_old).astype(float)

# 缩放比例（保护 old_range<=0 或 NaN）
scale = np.where((~old_range.isna()) & (old_range > 0), new_range / old_range, np.nan)

# 线性映射
result_mapped = T_LOW + (vals - low_old) * scale

# 缺失/异常保护
result_mapped = np.where(vals.isna() | np.isnan(scale), np.nan, result_mapped)

# 可选：裁剪到统一阈值范围
clip_to_new_range = False  # 是否把新结果裁剪到 [0, 100]
if clip_to_new_range:
    result_mapped = np.clip(result_mapped, 0, 100)

# 写入新列
df["结果_统一口径"] = result_mapped

# ---- （可选）导出 ----
if save_recalc_excel:
    # 确保目录存在
    os.makedirs(os.path.dirname(recalc_excel_path), exist_ok=True)
    df.to_excel(recalc_excel_path, index=False, engine="openpyxl")
    print(f"📄 已导出（含 T_low_unified / T_high_unified / 结果_统一口径）：{recalc_excel_path}")

# ---- 按工位绘图（Y 轴使用“结果_统一映射”）----
COLOR_MAP = {"中": "red", "优": "green", "良": "gold"}
stations = df[station_col].dropna().unique()

for station in stations:
    station_data = df[df[station_col] == station].copy().sort_values(by=time_col)
    colors = station_data["等级"].map(COLOR_MAP).fillna("gray")

    plt.figure(figsize=(11, 6.5))
    ax = plt.gca()

    # 浅灰线连点看趋势（基于映射后的结果）
    ax.plot(station_data[time_col], station_data["结果_统一口径"], color="#CFCFCF", linewidth=1, zorder=1)

    # 彩色散点
    ax.scatter(station_data[time_col], station_data["结果_统一口径"], c=colors, edgecolor="k", s=50, zorder=2)

    # 统一阈值参考线
    ax.axhline(T_LOW,  color="red",   linestyle="--", linewidth=1.2, label=f"T_low={T_LOW:.3f}")
    ax.axhline(T_HIGH, color="green", linestyle="--", linewidth=1.2, label=f"T_high={T_HIGH:.3f}")

    ax.set_title(f"station {station} result", fontsize=13)
    ax.set_xlabel("date")
    ax.set_ylabel("result (mapped)")
    plt.xticks(rotation=45)
    ax.grid(True, linestyle="--", alpha=0.4)
    ax.legend()
    plt.tight_layout()

    # 保存图表
    # 清洗工位名，避免 Windows 非法字符
    def safe_name(s):
        s = str(s)
        # 替换 Windows 不允许的字符： \ / : * ? " < > |
        s = re.sub(r'[\\\\/:*?"<>|]+', "_", s)
        # 可选：去掉首尾空格，限制长度
        s = s.strip()
        return s[:150]  # 防止过长路径问题

    safe_station = safe_name(station)

    # 生成安全文件名并保存
    fname = f"工位_{safe_station}.png"
    plt.savefig(os.path.join(output_dir, fname), dpi=150)
    plt.close()

print(f"✅ 图表已生成，保存在 {output_dir} 文件夹中")
print(f"👉 统一阈值：T_low={T_LOW:.3f}, T_high={T_HIGH:.3f}")