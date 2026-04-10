# Box_Plot
开发一个读取excel文件，并自动形成box plot的工具

## 简介

当前版本使用 Tkinter 作为桌面窗口 UI，不再依赖浏览器页面。

## 启动方式

1. 安装依赖：`pip install -r requirements.txt`
2. 启动应用：`python app.py`
3. 程序会直接打开桌面窗口

## 功能

- 支持导入多个 Excel 文件，并分别设置阶段名称与 Sheet。
- 支持两种数据布局：每列一个尺寸、每行一个尺寸。
- 支持预览前 20 行数据，并提取尺寸名称。
- 支持箱线图绘制、规格线显示、偏差显示和散点叠加。
- 支持显示统计摘要并导出 PNG 图片。
- 支持将当前配置保存为 JSON，并在后续加载复用。
- 启动时会自动读取默认配置文件 `box_plot_config.json`（若存在）。

## CPK 模板读取规则

- 在文件配置窗口可点击“套用 CPK 模板”。
- Sheet 由你手动选择（用于读取对应 CPK sheet）。
- 固定行映射如下（Excel 1-based）：
	- 第 10 行：dimension name
	- 第 11 行：nominal
	- 第 12 行：upper tolerance
	- 第 13 行：lower tolerance
	- 第 14 行：是否 GD&T
	- 第 79~111 行：测量数据
- GD&T 行中若标记为 `1 / Y / YES / TRUE / GD&T / 是`，按单边公差计算：
	- 优先使用 upper tolerance 计算 USL；若无 upper，则使用 lower tolerance 计算 LSL。
- 非 GD&T 尺寸按双边公差计算（同时考虑 upper/lower tolerance）。
- Cpk 支持双边和单边：
	- 双边：`min(Cpu, Cpl)`
	- 单边：仅按存在的一侧规格计算（Cpu 或 Cpl）

## 说明

- Tkinter 通常随 Windows 版 Python 一起安装，无需单独安装。
- Web 前端文件已移除，当前仅保留 Tkinter 桌面应用入口。

## 配置文件复用

1. 先完成一次文件配置（例如点击“套用 CPK 模板”并提取尺寸）。
2. 在主界面点击“保存配置”，导出为 `.json` 文件。
3. 下次使用时点击“加载配置”，即可把配置应用到当前文件，并设为默认模板。
4. 新增 Excel 文件时会自动尝试套用默认模板，减少重复配置。
5. 配置文件支持 `preferred_sheet` 字段，加载后会自动选中对应 sheet（例如 `CPK`）。
