import config
import pandas as pd

def Combined_rework_costs(file_path):
    # 读取各业务表，分别包含质量事件、产出数据、返工信息和过程结果
    df1 = pd.read_excel(file_path, sheet_name="pdca_incident_quality_info", dtype={"id": str, "process_result_id": str})
    df2 = pd.read_excel(file_path, sheet_name="pdca_biq_rework", dtype={"quality_info_id": str} )

    # 若存在调试数据则排除，避免脏数据影响统计
    if "debug_status" in df1.columns:
        df1 = df1[df1["debug_status"] != 1] 

    # 去除列名首尾空格，保证跨表字段匹配
    df1.columns = df1.columns.astype(str).str.strip()
    df2.columns = df2.columns.astype(str).str.strip()

    # 检查关键字段是否存在，缺失则直接中断
    need_df1_cols = ["id", "defect_number", "process_result_id", "date", "shift_id", "line_id", "area_id"]
    need_df2_cols = ["quality_info_id", '返工检测成本', "返工总成本"]
    missing1 = [c for c in need_df1_cols if c not in df1.columns]
    missing2 = [c for c in need_df2_cols if c not in df2.columns]
    if missing1:
        raise KeyError(f"df1 缺少列: {missing1}")
    if missing2:
        raise KeyError(f"df2 缺少列: {missing2}")
    df1 = df1[need_df1_cols].copy()
    df2 = df2[need_df2_cols].copy() 

    # 删除 process_result_id 为空的行（包括 NaN 和空字符串）
    df1 = df1[df1["process_result_id"].notna() & (df1["process_result_id"].astype(str).str.strip() != "")]
    # 根据map，替换id为对应的真实意义
    df1["process_result_name"] = df1["process_result_id"].map(config.process_result_map)
    df1["shift_name"] = df1["shift_id"].map(config.shift_name_map)
    df1["line_name"] = df1["line_id"].map(config.line_area_map)
    df1["area_name"] = df1["area_id"].map(config.station_name_map)
    df1['line_area_name'] = df1['line_name'].str.cat(df1['area_name'], sep=' ')

    df2["返工总成本"] = df2["返工总成本"] - df2["返工检测成本"]

    # 按列值匹配合并
    merged_df = pd.merge(
        df1, 
        df2, 
        left_on="id", 
        right_on="quality_info_id", 
        how="left"
    )

    return merged_df
    
def Real_output(file_path):
    # 读取各业务表，分别包含质量事件、产出数据、返工信息和过程结果
    df = pd.read_excel(file_path, sheet_name="oee_sun_shi_manager")

    # 去除列名首尾空格，保证跨表字段匹配
    df.columns = df.columns.astype(str).str.strip()

    # 检查关键字段是否存在，缺失则直接中断
    need_df_cols = ["date", "shift_id", "line_id", "regions_id", "real_out_put"]
    missing = [c for c in need_df_cols if c not in df.columns]
    if missing:
        raise KeyError(f"df 缺少列: {missing}")
    df = df[need_df_cols].copy()

    # 根据map，替换id为对应的真实意义
    df["shift_name"] = df["shift_id"].map(config.shift_name_map)
    df["line_name"] = df["line_id"].map(config.line_area_map)
    df["area_name"] = df["regions_id"].map(config.station_name_map)
    df['line_area_name'] = df['line_name'].str.cat(df['area_name'], sep=' ')

    return df  

if __name__ == "__main__":
    merged_df = Combined_rework_costs('./data/质量数据929.xlsx')
    merged_file = "./result/merged_929.xlsx"
    merged_df.to_excel(merged_file, index=False)

    output_df = Real_output('./data/质量数据929.xlsx')
    output_file = "./result/output_929.xlsx"
    output_df.to_excel(output_file, index=False)