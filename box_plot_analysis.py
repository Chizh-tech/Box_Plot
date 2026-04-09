"""
box_plot_analysis.py
--------------------
尺寸分析工具：读取 Excel 文件中各阶段的尺寸数据，生成对比箱线图。

Excel 文件格式约定
------------------
- 每张 Sheet 对应一个生产阶段（例如：毛坯、粗加工、精加工…）
- 每列对应一个尺寸特征（例如：尺寸A (mm)）
- 每行对应一次测量记录

运行方式
--------
    python box_plot_analysis.py                          # 使用默认文件 sample_data.xlsx
    python box_plot_analysis.py --file my_data.xlsx      # 指定 Excel 文件
    python box_plot_analysis.py --file data.xlsx --dim "尺寸A (mm)"   # 仅分析指定维度
    python box_plot_analysis.py --file data.xlsx --out results/       # 输出到指定目录
"""

import argparse
import os
import sys
from pathlib import Path
from typing import Optional

import matplotlib
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import numpy as np
import pandas as pd
import seaborn as sns

# 支持中文字体
matplotlib.rcParams["font.sans-serif"] = [
    "PingFang SC",
    "Microsoft YaHei",
    "SimHei",
    "DejaVu Sans",
]
matplotlib.rcParams["axes.unicode_minus"] = False

# ──────────────────────────────────────────────────────────────────────────────
# 数据读取
# ──────────────────────────────────────────────────────────────────────────────


def load_data(filepath: str) -> dict[str, pd.DataFrame]:
    """
    读取 Excel 文件，每张 Sheet 作为一个阶段。

    Returns
    -------
    dict[stage_name, DataFrame]
        键为阶段名称（Sheet 名），值为该阶段的测量数据。
    """
    path = Path(filepath)
    if not path.exists():
        sys.exit(f"错误：找不到文件 '{filepath}'")
    if path.suffix.lower() not in {".xlsx", ".xls"}:
        sys.exit(f"错误：不支持的文件格式 '{path.suffix}'，请使用 .xlsx 或 .xls")

    xls = pd.ExcelFile(filepath, engine="openpyxl")
    stages: dict[str, pd.DataFrame] = {}
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        # 只保留数值列
        numeric_df = df.select_dtypes(include="number")
        if numeric_df.empty:
            print(f"警告：Sheet '{sheet}' 中未找到数值列，已跳过。")
            continue
        stages[sheet] = numeric_df

    if not stages:
        sys.exit("错误：Excel 文件中没有可用的数值数据。")
    return stages


# ──────────────────────────────────────────────────────────────────────────────
# 统计摘要
# ──────────────────────────────────────────────────────────────────────────────


def compute_summary(stages: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """计算各阶段、各维度的描述性统计。"""
    records = []
    for stage, df in stages.items():
        for col in df.columns:
            s = df[col].dropna()
            records.append(
                {
                    "阶段": stage,
                    "维度": col,
                    "样本数": len(s),
                    "均值": round(s.mean(), 4),
                    "标准差": round(s.std(), 4),
                    "最小值": round(s.min(), 4),
                    "Q1": round(s.quantile(0.25), 4),
                    "中位数": round(s.median(), 4),
                    "Q3": round(s.quantile(0.75), 4),
                    "最大值": round(s.max(), 4),
                    "IQR": round(s.quantile(0.75) - s.quantile(0.25), 4),
                    "Cpk估算": _cpk(s),
                }
            )
    return pd.DataFrame(records)


def _cpk(series: pd.Series, usl_sigma: float = 3.0) -> Optional[float]:
    """简单 Cpk 估算（仅作参考，假设规格限为均值 ± usl_sigma × std）。"""
    std = series.std()
    if std == 0:
        return None
    mean = series.mean()
    usl = mean + usl_sigma * std
    lsl = mean - usl_sigma * std
    cpu = (usl - mean) / (3 * std)
    cpl = (mean - lsl) / (3 * std)
    return round(min(cpu, cpl), 3)


# ──────────────────────────────────────────────────────────────────────────────
# 绘图
# ──────────────────────────────────────────────────────────────────────────────


def _build_long_df(
    stages: dict[str, pd.DataFrame], dimension: str
) -> pd.DataFrame:
    """将各阶段指定维度的数据合并为长格式 DataFrame。"""
    frames = []
    for stage, df in stages.items():
        if dimension not in df.columns:
            continue
        tmp = pd.DataFrame({"阶段": stage, "测量值": df[dimension].dropna().values})
        frames.append(tmp)
    if not frames:
        raise ValueError(f"所有阶段中均未找到维度 '{dimension}'")
    result = pd.concat(frames, ignore_index=True)
    # 保持 Sheet 顺序
    result["阶段"] = pd.Categorical(result["阶段"], categories=list(stages.keys()), ordered=True)
    return result


def plot_boxplot(
    stages: dict[str, pd.DataFrame],
    dimension: str,
    output_dir: str = ".",
) -> str:
    """
    为指定维度绘制各阶段对比箱线图并保存。

    Returns
    -------
    str  保存路径
    """
    long_df = _build_long_df(stages, dimension)
    stage_order = long_df["阶段"].cat.categories.tolist()
    n_stages = len(stage_order)

    palette = sns.color_palette("Set2", n_stages)

    fig, ax = plt.subplots(figsize=(max(8, n_stages * 1.6), 6))

    sns.boxplot(
        data=long_df,
        x="阶段",
        y="测量值",
        hue="阶段",
        order=stage_order,
        palette=palette,
        width=0.5,
        linewidth=1.5,
        flierprops={"marker": "o", "markersize": 4, "alpha": 0.5},
        legend=False,
        ax=ax,
    )

    # 叠加散点（抖动）
    sns.stripplot(
        data=long_df,
        x="阶段",
        y="测量值",
        order=stage_order,
        color="black",
        size=3,
        alpha=0.3,
        jitter=True,
        ax=ax,
    )

    # 标注均值
    for i, stage in enumerate(stage_order):
        vals = long_df.loc[long_df["阶段"] == stage, "测量值"]
        mean_val = vals.mean()
        ax.plot(i, mean_val, marker="D", color="red", markersize=6, zorder=5)
        ax.text(
            i,
            mean_val,
            f" {mean_val:.3f}",
            va="bottom",
            ha="left",
            fontsize=8,
            color="red",
        )

    ax.set_title(f"各阶段尺寸对比箱线图\n{dimension}", fontsize=14, pad=12)
    ax.set_xlabel("生产阶段", fontsize=12)
    ax.set_ylabel(dimension, fontsize=12)
    ax.yaxis.grid(True, linestyle="--", alpha=0.7)
    ax.set_axisbelow(True)

    # 红色菱形 = 均值 图例
    legend_elements = [
        Line2D([0], [0], marker="D", color="w", markerfacecolor="red",
               markersize=8, label="均值"),
    ]
    ax.legend(handles=legend_elements, loc="upper right", fontsize=9)

    plt.tight_layout()

    safe_name = dimension.replace("/", "_").replace(" ", "_").replace("(", "").replace(")", "")
    os.makedirs(output_dir, exist_ok=True)
    save_path = os.path.join(output_dir, f"boxplot_{safe_name}.png")
    fig.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return save_path


def plot_summary_heatmap(
    summary_df: pd.DataFrame,
    metric: str = "标准差",
    output_dir: str = ".",
) -> str:
    """绘制各阶段 × 各维度的指标热力图。"""
    pivot = summary_df.pivot(index="维度", columns="阶段", values=metric)

    fig, ax = plt.subplots(figsize=(max(6, len(pivot.columns) * 1.5), max(4, len(pivot) * 1.0)))
    sns.heatmap(
        pivot,
        annot=True,
        fmt=".4f",
        cmap="YlOrRd",
        linewidths=0.5,
        ax=ax,
    )
    ax.set_title(f"各阶段各维度 {metric} 热力图", fontsize=13)
    plt.tight_layout()

    os.makedirs(output_dir, exist_ok=True)
    save_path = os.path.join(output_dir, f"heatmap_{metric}.png")
    fig.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return save_path


# ──────────────────────────────────────────────────────────────────────────────
# CLI 入口
# ──────────────────────────────────────────────────────────────────────────────


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="尺寸分析工具：从 Excel 读取各阶段数据并生成对比箱线图",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--file", "-f",
        default="sample_data.xlsx",
        help="Excel 数据文件路径（默认：sample_data.xlsx）",
    )
    parser.add_argument(
        "--dim", "-d",
        default=None,
        help="仅分析指定维度（列名），不指定则分析所有数值列",
    )
    parser.add_argument(
        "--out", "-o",
        default="output",
        help="图表输出目录（默认：output/）",
    )
    parser.add_argument(
        "--no-heatmap",
        action="store_true",
        help="不生成热力图",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    print(f"\n{'='*50}")
    print("  尺寸分析工具  |  Box Plot Analysis Tool")
    print(f"{'='*50}")
    print(f"数据文件：{args.file}")

    # 1. 读取数据
    stages = load_data(args.file)
    print(f"已读取 {len(stages)} 个阶段：{', '.join(stages.keys())}")

    # 获取所有维度
    all_dims: list[str] = []
    for df in stages.values():
        for col in df.columns:
            if col not in all_dims:
                all_dims.append(col)

    dims_to_plot = [args.dim] if args.dim else all_dims
    print(f"分析维度：{', '.join(dims_to_plot)}")

    # 2. 统计摘要
    summary = compute_summary(stages)
    summary_path = os.path.join(args.out, "summary.csv")
    os.makedirs(args.out, exist_ok=True)
    summary.to_csv(summary_path, index=False, encoding="utf-8-sig")
    print(f"\n统计摘要已保存：{summary_path}")
    print(summary.to_string(index=False))

    # 3. 箱线图（每个维度一张）
    print(f"\n生成箱线图...")
    for dim in dims_to_plot:
        try:
            path = plot_boxplot(stages, dim, output_dir=args.out)
            print(f"  ✓ {dim} → {path}")
        except ValueError as e:
            print(f"  ✗ {e}")

    # 4. 热力图
    if not args.no_heatmap:
        for metric in ["标准差", "均值"]:
            path = plot_summary_heatmap(summary, metric=metric, output_dir=args.out)
            print(f"  ✓ 热力图（{metric}）→ {path}")

    print(f"\n所有图表已保存至目录：{args.out}/")
    print("分析完成。\n")


if __name__ == "__main__":
    main()
