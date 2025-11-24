import pandas as pd
import numpy as np
import config

def calculate_ema(theta, beta, bias_correction=True):
    """
    计算指数加权移动平均(EMA)，支持偏差修正
    """
    theta = np.array(theta, dtype=np.float64)
    n = len(theta)
    v = np.zeros(n)  # 初始化EMA序列，v0=0（冷启动初始值）
    # 计算原始EMA
    for t in range(n):
        if t == 0:
            v[t] = (1 - beta) * theta[t]  # 算出来的v[0]其实是v1，因为v0=0，其实是 beta * 0 + (1 - beta) * theta[t]
        else:
            v[t] = beta * v[t-1] + (1 - beta) * theta[t]
    # 偏差修正：v_corrected = v[t] / (1 - beta ^ t)
    if bias_correction:
        t_array = np.arange(1, n+1)  # t从1到n，因为v[t]数组其实是从1开始的
        v_corrected = v / (1 - np.power(beta, t_array))
        return v, v_corrected
    else:
        return v, None
 
def calculate_ema_std(theta, beta, bias_correction=True):
    """
    计算指数加权标准差(EMA标准差)
    """
    theta = np.array(theta, dtype=np.float64)
    # 计算数据的EMA均值和数据平方的EMA均值
    _, mu = calculate_ema(theta, beta, bias_correction)
    _, nu = calculate_ema(theta**2, beta, bias_correction)
    # 计算方差并开方得到标准差
    variance = nu - mu**2
    # 确保方差非负（避免数值计算误差导致的微小负值）
    variance = np.maximum(variance, 0)
    std_values = np.sqrt(variance)
    return std_values
 
def ema_based_normalization(theta, beta, bias_correction=True):
    """
    使用EMA均值和EMA标准差进行数据标准化
    标准化公式: z_t = (theta_t - mu_t) / sigma_t
    """
    theta = np.array(theta, dtype=np.float64)
    # 计算EMA均值
    _, mu = calculate_ema(theta, beta, bias_correction)
    # 计算EMA标准差
    sigma = calculate_ema_std(theta, beta, bias_correction)
    # 避免除以零（添加微小值）
    sigma = np.maximum(sigma, 1e-6)
    data = (theta - mu)
    data[np.abs(data) < 1e-6] = 0.0
    # 标准化
    normalized_data = data / sigma
    return normalized_data, mu, sigma

def norma(theta, log_c, log_d):
    """
    对数缩放 + 归一化
    """
    # 对数变换，避免 log(0) 的问题
    log_theta = np.log(theta + log_c) / np.log(log_d) # c 可以是均值或一个固定值，比如 10

    norm_min, norm_max = (0, 1)
    global_min = np.min(log_theta)
    global_max = np.max(log_theta)
    
    # 检查是否已经在 [0, 1] 区间
    # if global_min >= norm_min and global_max <= norm_max:
    #     return theta  # 已经在 [0, 1]，直接返回

    # 处理特殊情况：如果数组所有元素几乎相等（避免除以0）
    if np.isclose(global_min, global_max, atol=1e-6):
        scaled_data = np.full_like(log_theta, norm_min)
    else:
        # 常规归一化：用全局范围缩放所有元素到[0,1]
        scaled_data = (log_theta - global_min) / (global_max - global_min) * (norm_max - norm_min) + norm_min
    return scaled_data

def Normaliz(df, beta, log_c, log_d):
    # Sort by 更新时间
    df = df.sort_values(by="更新时间")

    # Initialize dimensions
    x_values = sorted(df["工位"].unique())  # 工位序号从小到大
    z_values = sorted(df["更新时间"].unique())  # 时间从小到大

    # 指标列
    indicators = ["检验成本", "不合格率", "返工成本", "报废成本"]

    # Create empty array with shape (x=9, y=4, z=n)
    array_3d = np.empty((len(x_values), len(indicators), len(z_values)),dtype=np.float64)
    norm_array_3d = np.empty((len(x_values), len(indicators), len(z_values)),dtype=np.float64)
    scaled_array_3d = np.empty((len(x_values), len(indicators), len(z_values)),dtype=np.float64)
    print("3D array shape:", array_3d.shape)

    # Fill the array
    for z_idx, time in enumerate(z_values):
        df_time = df[df["更新时间"] == time]
        for x_idx, station in enumerate(x_values):
            df_entry = df_time[df_time["工位"] == station]
            if not df_entry.empty:
                array_3d[x_idx, :, z_idx] = df_entry[indicators].values[0]
            else:
                array_3d[x_idx, :, z_idx] = np.nan  # Fill with NaN if missing
    
    for x_idx, station in enumerate(x_values):
            for i, indicator in enumerate(indicators):
                normalized_data, mu, sigma = ema_based_normalization(array_3d[x_idx, i, :], beta)
                norm_array_3d[x_idx, i, :] = mu

    for i, indicator in enumerate(indicators):
        for z_idx, time in enumerate(z_values):
            # scaled_data = norm_array_3d[:, i, z_idx]
            scaled_data = norma(norm_array_3d[:, i, z_idx], log_c, log_d)
            scaled_array_3d[:, i, z_idx] = scaled_data
    
    # np.set_printoptions(suppress=True, threshold=np.inf, precision=6)
    # print(array_3d[:,1,0])
    # print(norm_array_3d[:,1,0])
    # print(scaled_array_3d[:,1,0])   
     
    rows = []
    # 遍历所有「工位+时间」组合，提取对应指标值
    for z_idx, time in enumerate(z_values):
        for x_idx, station in enumerate(x_values):
            # 当前工位+时间的所有指标值（第二维度）
            indicator_vals = scaled_array_3d[x_idx, :, z_idx]
            # 构建一行数据（键：列名，值：对应数据）
            row = {
                "更新时间": time,
                "工位": station,
                **dict(zip(indicators, indicator_vals))  # 指标与值一一对应
            }
            rows.append(row)

    # 转换为DataFrame并设置多索引（方便按工位/时间筛选）
    scaled_df = pd.DataFrame(rows)
    # scaled_df = scaled_df.set_index(["工位", "更新时间"])

    return scaled_df 

def main():
    # Load the Excel file
    df = pd.read_excel("./data/final_result.xlsx", engine="openpyxl")
    
    df["更新时间"] = df["date"]
    df["不合格率"] = 1 - df["合格率"]

    need_df_cols = ["工位", "更新时间", "检验成本", "不合格率", "返工成本", "报废成本"]
    df = df[need_df_cols].copy()

    results_df = Normaliz(df, config.beta, config.log_c, config.log_d)

    results_df.to_excel("./result/datastandard.xlsx", index=False)

    # 设置显示所有行（取消行数限制）
    pd.set_option('display.max_rows', None)
    # 设置显示所有列（取消列数限制）
    pd.set_option('display.max_columns', None)
    # 设置列宽（避免内容被截断）
    pd.set_option('display.max_colwidth', None)
    pd.set_option('display.float_format', '{:.6f}'.format)

    # 打印完整DataFrame
    print(results_df)

if __name__ == "__main__":
    main()