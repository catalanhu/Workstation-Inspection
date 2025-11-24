import pandas as pd

def split_and_sort_by_station(input_file, output_file):
    # 1. 读取Excel数据（确保表头正确，时间列解析为datetime类型）
    # 若列名不同，需修改header参数或后续列名（例如：header=0表示第一行为表头）
    df = pd.read_excel(
        input_file,
        # sheet_name="oee_sun_shi_manager",
        parse_dates=['date'],  # 将“时间”列解析为datetime类型，确保排序正确
        engine='openpyxl'
    )
    
    # 2. 定义班次排序优先级（白班→中班→夜班）
    shift_priority = {'白班': 0, '中班': 1, '夜班': 2}
    
    # 3. 按“工位”分组，遍历每个工位处理数据
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 遍历每个工位的分组数据
        for station, group in df.groupby('line_area_name'):
            # 4. 排序：先按“时间”升序，时间相同则按“班次”优先级排序
            # 为“班次”列映射优先级数值，用于排序
            group['班次优先级'] = group['shift_id'].map(shift_priority)
            # 按“时间”和“班次优先级”排序（排序后删除临时列）
            sorted_group = group.sort_values(
                by=['date', '班次优先级'],  # 排序依据：先时间，后班次优先级
                ascending=[True, True]      # 均为升序（时间从小到大，班次按白→中→夜）
            ).drop(columns=['班次优先级'])  # 删除临时列
            
            # 5. 将排序后的数据写入以工位命名的子表
            # 工作表名称最大31字符，且不能含特殊符号（假设工位名称合法）
            sorted_group.to_excel(
                writer,
                sheet_name=station,  # 子表名称=工位名称
                index=False          # 不写入索引列
            )
            print(f"已生成工位「{station}」的子表，共{len(sorted_group)}条数据")
    
    print(f"\n所有子表已保存至：{output_file}")

# 执行函数（替换为你的输入/输出文件路径）
if __name__ == "__main__":
    input_excel = "merged_result.xlsx"   # 原始Excel文件路径
    output_excel = "output2.xlsx" # 输出结果文件路径
    split_and_sort_by_station(input_excel, output_excel)