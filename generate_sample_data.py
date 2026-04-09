"""
generate_sample_data.py
-----------------------
生成示例 Excel 文件，用于演示尺寸分析工具。

每张 Sheet 代表一个生产阶段，每列代表一个尺寸特征，
每行代表一次测量记录。

运行方式：
    python generate_sample_data.py
"""

import numpy as np
import pandas as pd

np.random.seed(42)

# ---------- 各阶段尺寸数据配置（均值, 标准差, 样本数） ----------
STAGES = {
    "毛坯":       {"mean": 100.5, "std": 0.80, "n": 50},
    "粗加工":     {"mean": 100.2, "std": 0.45, "n": 50},
    "半精加工":   {"mean": 100.08, "std": 0.20, "n": 50},
    "精加工":     {"mean": 100.02, "std": 0.08, "n": 50},
    "检验":       {"mean": 100.01, "std": 0.05, "n": 50},
}

DIMENSIONS = ["尺寸A (mm)", "尺寸B (mm)", "尺寸C (mm)"]

OUTPUT_FILE = "sample_data.xlsx"


def generate_stage_data(cfg: dict) -> pd.DataFrame:
    """为一个阶段生成多维度测量数据。"""
    data = {}
    for i, dim in enumerate(DIMENSIONS):
        # 每个维度的目标值略有偏移，以体现实际差异
        offset = i * 0.3
        data[dim] = np.random.normal(
            loc=cfg["mean"] + offset,
            scale=cfg["std"],
            size=cfg["n"],
        ).round(4)
    return pd.DataFrame(data)


def main() -> None:
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        for stage, cfg in STAGES.items():
            df = generate_stage_data(cfg)
            df.to_excel(writer, sheet_name=stage, index=False)
    print(f"示例数据已生成：{OUTPUT_FILE}")
    print(f"包含阶段：{', '.join(STAGES.keys())}")
    print(f"包含维度：{', '.join(DIMENSIONS)}")


if __name__ == "__main__":
    main()
