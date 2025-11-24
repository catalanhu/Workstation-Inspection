import pandas as pd
from datetime import timedelta

def calculate_90days_sum_with_shift_order(input_file, output_file):
    # 1. 定义班次优先级（白班→中班→晚班）
    shift_priority = {'白班': 0, '中班': 1, '夜班': 2}  # 确保与实际班次名称一致
    
    # 2. 读取Excel文件并获取所有子表名称
    excel_file = pd.ExcelFile(input_file, engine='openpyxl')
    sheet_names = excel_file.sheet_names  # 每个子表对应一个工位

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet in sheet_names:
            # 3. 读取当前子表数据
            df = excel_file.parse(sheet)
            
            # 检查必要列是否存在（日期、班次、产量）
            required_cols = ['date', 'shift_id', 'real_out_put']
            if not set(required_cols).issubset(df.columns):
                print(f"❌ 子表「{sheet}」缺少必要列（{required_cols}），跳过处理")
                continue
            
            # 4. 解析日期列并处理班次排序
            df['date'] = pd.to_datetime(df['date'])  # 确保日期格式正确
            # 添加班次优先级列（用于排序）
            df['班次优先级'] = df['shift_id'].map(shift_priority)
            # 按“日期升序→班次优先级升序”排序（同一天内白班→中班→晚班）
            df = df.sort_values(by=['date', '班次优先级'], ascending=[True, True])
            
            # 5. 计算90天内real_out_put总和（基于排序后的数据）
            df = df.set_index('date')  # 日期设为索引，用于时间窗口
            df['90天real_out_put总和'] = df['real_out_put'].rolling('90D', closed='both').sum()
            df = df.reset_index()  # 重置索引
            
            # 6. 筛选“有完整90天数据”的记录（当前日期 ≥ 最早日期+90天）
            first_date = df['date'].min()
            valid_records = df[df['date'] >= first_date + timedelta(days=90)]
            
            # 7. 删除临时的“班次优先级”列，保留原始数据
            valid_records = valid_records.drop(columns=['班次优先级'], errors='ignore')
            
            # 8. 写入结果到对应子表
            valid_records.to_excel(writer, sheet_name=sheet, index=False)
            print(f"✅ 子表「{sheet}」处理完成，有效记录数：{len(valid_records)}")

    print(f"\n所有结果已保存至：{output_file}")

if __name__ == "__main__":
    input_excel = "output.xlsx"   # 替换为实际输入路径
    output_excel = "各子表_90天总和（含班次排序）.xlsx"  # 输出路径
    calculate_90days_sum_with_shift_order(input_excel, output_excel)