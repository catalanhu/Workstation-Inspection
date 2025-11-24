import pandas as pd
import matplotlib.pyplot as plt
import os

# 读取 Excel 文件
file_path = "./result/test_result.xlsx"  # 替换为你的文件路径
df = pd.read_excel(file_path)

# 确认列名
time_col = df.columns[0]      # 第一列：时间
station_col = df.columns[1]   # 第二列：工位
result_col = "结果"            # 替换为实际列名（截图里靠右的那列）

# 转换时间格式
df[time_col] = pd.to_datetime(df[time_col])

# 创建输出目录
output_dir = "./charts"
os.makedirs(output_dir, exist_ok=True)

# 按工位分组
stations = df[station_col].unique()

for station in stations:
    station_data = df[df[station_col] == station].sort_values(by=time_col)
    
    plt.figure(figsize=(10, 6))
    plt.plot(station_data[time_col], station_data[result_col], marker='o')
    
    plt.title(f"station {station} result")
    plt.xlabel("date")
    plt.ylabel("result")
    plt.xticks(rotation=45)
    plt.grid(True)
    plt.tight_layout()
    
    # 保存图表
    plt.savefig(os.path.join(output_dir, f"工位_{station}.png"))
    plt.close()

print(f"✅ 图表已生成，保存在 {output_dir} 文件夹中")