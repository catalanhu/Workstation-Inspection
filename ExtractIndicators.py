import pandas as pd
import ExtractData as ED
import StationSegmentation as SS
import config

def merge_summary_data(summary_data, summary_output):
    merged_result = {}

    # 找出两个字典中共同的工位名称
    common_stations = set(summary_data.keys()).intersection(set(summary_output.keys()))

    for station in common_stations:
        df1 = summary_data[station]      # 缺陷统计
        df2 = summary_output[station]    # 实际产出统计

        # 合并：根据 date 对齐
        merged_df = pd.merge(
            df1,
            df2,
            on='date',
            how='outer'
        )
        # 按日期排序
        merged_df = merged_df.sort_values(by='date').reset_index(drop=True)
        merged_result[station] = merged_df

    return merged_result

def compute_quality_metrics(merged_result, window_days=90):
    result_with_metrics = {}
    for station, df in merged_result.items():
        # 用最近有数据的值填充缺失（前向填充）
        defect_cols = ['总缺陷数', '返工总数', '报废总数', '其他总数', '返工检测成本', '返工总成本']
        # 如果仍有 NaN（比如开头没有数据），填 0
        df[defect_cols] = df[defect_cols].fillna(0)
        
        # 计算过去 90 天滚动累计值
        for col in defect_cols:
            df[col] = df.set_index('date')[col].rolling(f'{window_days}D').sum().values

        # 删除没有产量的记录
        df = df[df['daily_total_output'].notna() & (df['daily_total_output'] != 0)].copy()

        # 计算指标
        df['检验成本'] = df['返工检测成本']
        df['合格率'] = 1 - (df['总缺陷数'] / df['daily_total_output'])
        df['返工成本'] = df['返工总成本']
        df['报废成本'] = df['报废总数'] * 1    # 报废单价目前设为1

        result_with_metrics[station] = df.reset_index(drop=True)

    return result_with_metrics

def consolidate_metrics(result_with_metrics):
    # 收集每个工位的日期集合
    station_dates = {station: set(df['date']) for station, df in result_with_metrics.items()}

    # 找出所有工位的公共日期
    common_dates = set.intersection(*station_dates.values())

    # 按工位名称排序筛选，过滤掉不在公共日期中的数据
    station_order = sorted(result_with_metrics.keys())  # fixed order
    filtered_dfs = []
    for station in station_order:
        df = result_with_metrics[station]
        df_filtered = df[df['date'].isin(common_dates)].copy()
        df_filtered['station'] = station
        filtered_dfs.append(df_filtered)

    # 合并过滤后的数据
    combined_df = pd.concat(filtered_dfs, ignore_index=True)

    # 按日期和工位顺序排序
    combined_df['date'] = pd.to_datetime(combined_df['date'])
    combined_df['station_order'] = combined_df['station'].map({name: i for i, name in enumerate(station_order)})
    combined_df = combined_df.sort_values(by=['date', 'station_order']).drop(columns=['station_order']).reset_index(drop=True)

    # 选取需要的数据
    final_df = combined_df[['date', 'station', '检验成本', '合格率', '返工成本', '报废成本']]
    
    final_df = final_df.copy()
    final_df.rename(columns={"station": "工位"}, inplace=True)
    final_df["工位"] = final_df["工位"].map(config.station_map)

    return final_df

if __name__ == "__main__":  
    merged_df = ED.Combined_rework_costs("./data/质量数据929.xlsx")
    station_data = SS.split_station(merged_df)
    summary_data = SS.shift_summary(station_data) 
    daily_summary_data = SS.daily_total(summary_data)

    output_df = ED.Real_output("./data/质量数据929.xlsx")
    station_output = SS.split_station_output(output_df)
    summary_output = SS.shift_summary_output(station_output) 
    daily_output = SS.daily_total_output(summary_output) 

    merged_result = merge_summary_data(daily_summary_data, daily_output)
    output_file = "./result/Combined_Summary.xlsx"
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for station, df in merged_result.items():
            df.to_excel(writer, sheet_name=station, index=False)
            print(f"✅ 已合并工位: {station}, 行数: {len(df)}")

    result_with_metrics = compute_quality_metrics(merged_result)
    result = "./result/result.xlsx"
    
    with pd.ExcelWriter(result, engine='openpyxl') as writer:
        for station, df in result_with_metrics.items():
            df.to_excel(writer, sheet_name=station, index=False)
            print(f"✅ 已生成结果: {station}, 行数: {len(df)}")

    final_df = consolidate_metrics(result_with_metrics)
    final_result = "./result/final_result.xlsx"
    final_df.to_excel(final_result, index=False)
