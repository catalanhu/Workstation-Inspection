import pandas as pd
import ExtractData as ED

def split_station(merged_df):    
    # 定义班次排序优先级（白班→中班→夜班）
    shift_priority = {'白班': 0, '中班': 1, '夜班': 2}

    # 创建一个字典用于存储每个工位的DataFrame
    station_data = {}
    
    # 遍历每个工位的分组数据
    for station, group in merged_df.groupby('line_area_name'):
        # 排序：先按“时间”升序，时间相同则按“班次”优先级排序
        group['班次优先级'] = group['shift_name'].map(shift_priority)
        # 按“时间”和“班次优先级”排序（排序后删除临时列）
        sorted_group = group.sort_values(
            by=['date', '班次优先级'],  # 排序依据：先时间，后班次优先级
            ascending=[True, True]      # 均为升序（时间从小到大，班次按白→中→夜）
        ).drop(columns=['班次优先级'])  # 删除临时列
        
        # 只保留指定的列
        selected_columns = ['date', 'shift_name', 'process_result_name', 'defect_number', '返工检测成本', '返工总成本']
        filtered_group = sorted_group[selected_columns].reset_index(drop=True)
                    
        # 保存到字典中
        station_data[station] = filtered_group

    return station_data

def shift_summary(station_data):
    summary_data = {}
    for station, df in station_data.items():
        grouped = df.groupby(['date', 'shift_name'])
        result = []
        for (date, shift), group in grouped:
            total_defects = group['defect_number'].sum()
            rework_total = group[group['process_result_name'] == '返工']['defect_number'].sum()
            scrap_total = group[group['process_result_name'] == '报废']['defect_number'].sum()
            other_total = group[~group['process_result_name'].isin(['返工', '报废'])]['defect_number'].sum()
            check_cost_total = group[group['process_result_name'] == '返工']['返工检测成本'].sum()
            rework_cost_total = group[group['process_result_name'] == '返工']['返工总成本'].sum()
            result.append({
                'date': date,
                'shift_name': shift,
                '总缺陷数': total_defects,
                '返工总数': rework_total,
                '报废总数': scrap_total,
                '其他总数': other_total,
                '返工检测成本': check_cost_total,
                '返工总成本': rework_cost_total
            })
        summary_df = pd.DataFrame(result)

        shift_priority = {'白班': 0, '中班': 1, '夜班': 2}
        summary_df['班次优先级'] = summary_df['shift_name'].map(shift_priority)
        summary_df = summary_df.sort_values(by=['date', '班次优先级']).drop(columns=['班次优先级']).reset_index(drop=True)

        summary_data[station] = summary_df
    return summary_data

def daily_total(summary_data, window_days=1):  
    daily_summary_data = {}
    for station, df in summary_data.items():
        # 按日期汇总各项指标
        daily_df = df.groupby('date').agg({
            '总缺陷数': 'sum',
            '返工总数': 'sum',
            '报废总数': 'sum',
            '其他总数': 'sum',
            '返工检测成本': 'sum',
            '返工总成本': 'sum'
        }).reset_index()

        
        # 确保 date 是 datetime 类型
        daily_df['date'] = pd.to_datetime(daily_df['date'])
        daily_df = daily_df.sort_values(by='date').set_index('date')

        # 基于时间窗口滚动累计
        daily_df = daily_df.rolling(f'{window_days}D', min_periods=1).sum()

        # 重置索引
        daily_df = daily_df.reset_index()

        # 去掉不足 window_days 的数据
        start_date = daily_df['date'].min()
        daily_df = daily_df[daily_df['date'] >= start_date + pd.Timedelta(days=window_days)]

        daily_summary_data[station] = daily_df
    return daily_summary_data


def split_station_output(output_df):
    # 定义班次排序优先级（白班→中班→夜班）
    shift_priority = {'白班': 0, '中班': 1, '夜班': 2}

    # 创建一个字典用于存储每个工位的DataFrame
    station_output = {}
    
    # 遍历每个工位的分组数据
    for station, group in output_df.groupby('line_area_name'):
        # 排序：先按“时间”升序，时间相同则按“班次”优先级排序
        group['班次优先级'] = group['shift_name'].map(shift_priority)
        # 按“时间”和“班次优先级”排序（排序后删除临时列）
        sorted_group = group.sort_values(
            by=['date', '班次优先级'],  # 排序依据：先时间，后班次优先级
            ascending=[True, True]      # 均为升序（时间从小到大，班次按白→中→夜）
        ).drop(columns=['班次优先级'])  # 删除临时列
        
        # 只保留指定的列
        selected_columns = ['date', 'shift_name', 'real_out_put']
        filtered_group = sorted_group[selected_columns].reset_index(drop=True)
                    
        # 保存到字典中
        station_output[station] = filtered_group

    return station_output

def shift_summary_output(station_output):
    summary_output = {}
    for station, df in station_output.items():
        grouped = df.groupby(['date', 'shift_name'])
        result = []
        for (date, shift), group in grouped:
            total_defects = group['real_out_put'].sum()
            result.append({
                'date': date,
                'shift_name': shift,
                'total_real_output': total_defects,
            })
        real_output_df = pd.DataFrame(result)

        shift_priority = {'白班': 0, '中班': 1, '夜班': 2}
        real_output_df['班次优先级'] = real_output_df['shift_name'].map(shift_priority)
        real_output_df = real_output_df.sort_values(by=['date', '班次优先级']).drop(columns=['班次优先级']).reset_index(drop=True)

        summary_output[station] = real_output_df
    return summary_output

def daily_total_output(summary_output, window_days=90):
    daily_output = {}
    for station, df in summary_output.items():
        # 按日期汇总所有班次的产量
        daily_totals = df.groupby('date')['total_real_output'].sum().reset_index()
        daily_totals = daily_totals.rename(columns={'total_real_output': 'daily_total_output'})
        
        # 确保 date 是 datetime 类型
        daily_totals['date'] = pd.to_datetime(daily_totals['date'])
        daily_totals = daily_totals.sort_values(by='date').set_index('date')

        # 基于时间窗口滚动累计
        daily_totals['daily_total_output'] = daily_totals['daily_total_output'].rolling(f'{window_days}D', min_periods=1).sum()

        # 重置索引
        daily_totals = daily_totals.reset_index()
   
        # 去掉不足 window_days 的数据
        start_date = daily_totals['date'].min()
        daily_totals = daily_totals[daily_totals['date'] >= start_date + pd.Timedelta(days=window_days)]

        daily_output[station] = daily_totals
    return daily_output

if __name__ == "__main__":
    merged_df = ED.Combined_rework_costs("./data/质量数据929.xlsx")
    merged_1 = "./result/Station_1.xlsx"
    merged_2 = "./result/Station_2.xlsx"
    merged_3 = "./result/Station_3.xlsx"
    
    output_df = ED.Real_output("./data/质量数据929.xlsx")
    output_1 = "./result/output_1.xlsx"
    output_2 = "./result/output_2.xlsx"
    output_3 = "./result/output_3.xlsx"

    # Step 1: 获取每个工位的已排序数据
    station_data = split_station(merged_df)
    with pd.ExcelWriter(merged_1, engine='openpyxl') as writer:
        for station, df in station_data.items():
            df.to_excel(writer, sheet_name=station, index=False)
            print(f"已生成工位「{station}」的子表，共{len(df)}条数据")

    # Step 2: 聚合每个工位每个班次的缺陷数据
    summary_data = shift_summary(station_data) 
    with pd.ExcelWriter(merged_2, engine='openpyxl') as writer:
            for station, df in summary_data.items():
                df.to_excel(writer, sheet_name=station, index=False)
                print(f"已生成工位「{station}」的统计子表，共{len(df)}条记录")

    # Step 3: 聚合每个工位每天的缺陷数据
    daily_summary_data = daily_total(summary_data)
    with pd.ExcelWriter(merged_3, engine='openpyxl') as writer:
            for station, df in daily_summary_data.items():
                df.to_excel(writer, sheet_name=station, index=False)
                print(f"已生成工位「{station}」的统计子表，共{len(df)}条记录")

    # Step 4: 获取每个工位的已排序产量
    station_output = split_station_output(output_df)
    with pd.ExcelWriter(output_1, engine='openpyxl') as writer:
        for station, df in station_output.items():
            df.to_excel(writer, sheet_name=station, index=False)
            print(f"已生成工位「{station}」的子表，共{len(df)}条数据")

    # Step 5: 聚合每个工位每个班次的产量
    summary_output = shift_summary_output(station_output) 
    with pd.ExcelWriter(output_2, engine='openpyxl') as writer:
            for station, df in summary_output.items():
                df.to_excel(writer, sheet_name=station, index=False)
                print(f"已生成工位「{station}」的统计子表，共{len(df)}条记录")

    # Step 6: 聚合每个工位每天的产量
    daily_output = daily_total_output(summary_output) 
    with pd.ExcelWriter(output_3, engine='openpyxl') as writer:
            for station, df in daily_output.items():
                df.to_excel(writer, sheet_name=station, index=False)
                print(f"已生成工位「{station}」的统计子表，共{len(df)}条记录")