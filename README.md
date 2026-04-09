# Box_Plot 尺寸分析工具

读取 Excel 文件中各生产阶段的尺寸测量数据，自动生成对比箱线图，辅助过程能力分析。

## 功能特性

- 📊 **多阶段箱线图**：每个尺寸维度生成一张对比图，展示各阶段分布差异
- 🔴 **均值标注**：图中标注各阶段均值，一目了然
- 🌡️ **热力图**：以热力图呈现各阶段 × 各维度的均值与标准差
- 📋 **统计摘要**：输出 CSV 文件，含均值、标准差、四分位数、IQR、Cpk 估算
- 🗂️ **灵活格式**：Excel 每张 Sheet 对应一个阶段，每列对应一个尺寸特征

## Excel 文件格式约定

| Sheet 名称 | 列名（尺寸特征） |
|-----------|---------------|
| 毛坯       | 尺寸A (mm)、尺寸B (mm)、… |
| 粗加工     | 尺寸A (mm)、尺寸B (mm)、… |
| 精加工     | …             |

- **每张 Sheet** = 一个生产阶段
- **每列** = 一个测量维度
- **每行** = 一次测量记录

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 生成示例数据（可选）

```bash
python generate_sample_data.py
# 输出：sample_data.xlsx（5 个阶段 × 3 个维度 × 50 条记录）
```

### 3. 运行分析

```bash
# 使用默认示例文件
python box_plot_analysis.py

# 指定自己的 Excel 文件
python box_plot_analysis.py --file my_data.xlsx

# 只分析某一维度
python box_plot_analysis.py --file my_data.xlsx --dim "尺寸A (mm)"

# 指定输出目录
python box_plot_analysis.py --file my_data.xlsx --out results/

# 不生成热力图
python box_plot_analysis.py --no-heatmap
```

### 4. 查看输出

```
output/
├── boxplot_尺寸A_mm.png   # 各阶段箱线图（每个维度一张）
├── boxplot_尺寸B_mm.png
├── boxplot_尺寸C_mm.png
├── heatmap_标准差.png      # 标准差热力图
├── heatmap_均值.png        # 均值热力图
└── summary.csv            # 统计摘要（含 Cpk 估算）
```

## 命令行参数

| 参数 | 简写 | 默认值 | 说明 |
|------|------|--------|------|
| `--file` | `-f` | `sample_data.xlsx` | Excel 数据文件路径 |
| `--dim` | `-d` | 全部 | 仅分析指定维度列名 |
| `--out` | `-o` | `output` | 图表输出目录 |
| `--no-heatmap` | — | 否 | 不生成热力图 |

## 文件结构

```
Box_Plot/
├── box_plot_analysis.py    # 主程序
├── generate_sample_data.py # 示例数据生成器
├── requirements.txt        # Python 依赖
├── sample_data.xlsx        # 示例数据（运行生成器后产生）
└── output/                 # 图表输出目录（运行分析后产生）
```
