import os
import re
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ================== 可配置区 ==================
file_path = "./result/test_result.xlsx"   # 输入Excel 文件路径
output_dir = "./charts_result_html"            # HTML 输出目录
os.makedirs(output_dir, exist_ok=True)
save_recalc_excel = False                  # 是否导出带映射结果与新等级的Excel
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

# 写入新列
df["结果_统一口径"] = result_mapped

# ---- （可选）导出 ----
if save_recalc_excel:
    os.makedirs(os.path.dirname(recalc_excel_path), exist_ok=True)
    df.to_excel(recalc_excel_path, index=False, engine="openpyxl")
    print(f"📄 已导出（含 T_low_unified / T_high_unified / 结果_统一口径）：{recalc_excel_path}")

# ---- 按工位生成交互式 HTML 图表 ----
stations = df[station_col].dropna().unique()

for station in stations:
    station_data = df[df[station_col] == station].copy().sort_values(by=time_col)

    # 创建 Plotly 散点图
    fig = px.scatter(
        station_data,
        x=time_col,
        y="结果_统一口径",
        color="等级" if "等级" in station_data.columns else None,
        hover_data={time_col: True, "结果_统一口径": ':.2f', "等级": True},
        title=f"station {station} result"
    )

    # 添加趋势线（灰色）
    fig.add_trace(go.Scatter(
        x=station_data[time_col],
        y=station_data["结果_统一口径"],
        mode='lines',
        line=dict(color='lightgray', width=1),
        showlegend=False
    ))

    # 添加阈值线
    fig.add_hline(y=T_LOW, line_dash="dash", line_color="red", annotation_text=f"T_low={T_LOW:.3f}")
    fig.add_hline(y=T_HIGH, line_dash="dash", line_color="green", annotation_text=f"T_high={T_HIGH:.3f}")

    # 保存 HTML 文件
    safe_station = re.sub(r'[\\/:*?"<>|]+', "_", str(station)).strip()[:150]
    fname = f"工位_{safe_station}.html"
    fig.write_html(os.path.join(output_dir, fname))

print(f"✅ 交互式图表已生成，保存在 {output_dir} 文件夹中（HTML 格式）")
print(f"👉 统一阈值：T_low={T_LOW:.3f}, T_high={T_HIGH:.3f}")