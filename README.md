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

## 说明

- Tkinter 通常随 Windows 版 Python 一起安装，无需单独安装。
- 旧的 `templates/` 和 `static/` 目录仍保留在仓库中，但桌面版启动时不会使用它们。
