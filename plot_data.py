import os
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib

# ----------------- 可配置区 -----------------
EXCEL_PATH = "./data/final_result.xlsx"
DATE_COL   = "date"
STATION_COL = "工位"
METRICS = ["检验成本", "合格率", "返工成本", "报废成本"]  # 4 个指标
OUTPUT_DIR = "./charts_data"
STATION_RANGE = list(range(1, 27))  # 固定显示工位 1~26
# ------------------------------------------

def setup_chinese_font():
    # 尝试让中文不乱码
    matplotlib.rcParams['axes.unicode_minus'] = False
    for f in ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS', 'Noto Sans CJK SC']:
        try:
            plt.rcParams['font.sans-serif'] = [f]
            return
        except Exception:
            continue

def safe_name(s: str, max_len=150):
    s = str(s)
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    return s.strip()[:max_len]

def pick_two_dates(df):
    # 唯一日期列表（排序）
    unique_dates = df[DATE_COL].dropna().dt.normalize().sort_values().unique()
    if len(unique_dates) < 2:
        raise ValueError("文件中可用日期少于 2 个，无法对比。")

    def fmt(d): return pd.to_datetime(d).strftime("%Y-%m-%d")

    print("请输入两个要对比的日期")
    print("可用日期：", ", ".join(fmt(d) for d in unique_dates))

    def ask(msg):
        while True:
            s = input(msg).strip()
            try:
                d = pd.to_datetime(s).normalize()
                if d in unique_dates:
                    return d
                else:
                    print("该日期不在数据中，请重新输入。")
            except Exception:
                print("格式错误，请输入 YYYY-MM-DD。")

    d1 = ask("第一个日期：")
    d2 = ask("第二个日期：")
    return d1, d2

def plot_metric(wide1, wide2, d1, d2, metric):
    # 柱状对比：两个日期并列展示 1..26 工位
    x = np.arange(len(STATION_RANGE))
    width = 0.38

    fig, ax = plt.subplots(figsize=(12, 5))
    ax.bar(x - width/2, wide1.values, width, label=pd.to_datetime(d1).strftime("%Y-%m-%d"),
           color="#1f77b4")
    ax.bar(x + width/2, wide2.values, width, label=pd.to_datetime(d2).strftime("%Y-%m-%d"),
           color="#ff7f0e")

    ax.set_xticks(x)
    ax.set_xticklabels(STATION_RANGE, rotation=0)
    ax.set_xlabel("stations")
    ax.set_ylabel(metric)
    title = f"{metric} (station 1~26)"
    ax.set_title(title)
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    ax.legend()

    # 在柱顶加数值（可按需保留 2 位小数）
    def annotate_bars(bars):
        for b in bars:
            h = b.get_height()
            ax.annotate(f"{h:.2f}" if np.isfinite(h) else "",
                        xy=(b.get_x() + b.get_width()/2, h if np.isfinite(h) else 0),
                        xytext=(0, 3), textcoords="offset points",
                        ha="center", va="bottom", fontsize=5.5)

    # 重新获取 patch（两拨柱子）
    bars1 = [p for p in ax.patches[:len(STATION_RANGE)]]
    bars2 = [p for p in ax.patches[len(STATION_RANGE):]]
    annotate_bars(bars1)
    annotate_bars(bars2)

    plt.tight_layout()
    return fig, ax

def main():
    setup_chinese_font()
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
    df[STATION_COL] = pd.to_numeric(df[STATION_COL], errors="coerce").astype("Int64")
    for m in METRICS:
        df[m] = pd.to_numeric(df[m], errors="coerce")

    d1, d2 = pick_two_dates(df)

    print(f"将对比的日期：{pd.to_datetime(d1).strftime('%Y-%m-%d')}  vs  {pd.to_datetime(d2).strftime('%Y-%m-%d')}")
    for metric in METRICS:
        wide1 = df[df[DATE_COL].dt.normalize() == d1].set_index(STATION_COL)[metric] 
        wide2 = df[df[DATE_COL].dt.normalize() == d2].set_index(STATION_COL)[metric]
        
        # 仅对非“合格率”指标取 log
        if metric != "合格率":
            wide1 = np.log(wide1 + 1) / np.log(10)
            wide2 = np.log(wide2 + 1) / np.log(10)

        fig, ax = plot_metric(wide1, wide2, d1, d2, metric)

        fname = f"{safe_name(metric)}_{pd.to_datetime(d1).strftime('%Y%m%d')}_VS_{pd.to_datetime(d2).strftime('%Y%m%d')}.png"
        out_path = os.path.join(OUTPUT_DIR, fname)
        plt.savefig(out_path, dpi=150)
        plt.close(fig)
        print(f"✅ 已保存：{out_path}")

    print(f"全部完成，图像保存在：{os.path.abspath(OUTPUT_DIR)}")

if __name__ == "__main__":
    main()