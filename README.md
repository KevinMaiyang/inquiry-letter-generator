# 询证函生成器

> 一个用于从 Excel 台账批量生成询证函的桌面应用程序，支持模板自定义、多工作表处理、格式保留。

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-lightgrey)

## 📌 功能特点

- **批量生成**：根据 Excel 台账每行数据，自动生成独立工作表的询证函
- **模板自定义**：可编辑回函地址、联系人、电话、邮箱、发函日期等字段
- **智能日期**：自动根据发函日期计算对应季度（如“2025年第3季度”）
- **格式保留**：完整保留模板中的合并单元格、列宽、行高、字体样式
- **持久化配置**：修改的模板字段会自动保存，下次启动自动加载
- **跨平台支持**：提供 Windows `.exe` 可执行文件

## 📂 使用说明

### 1. 准备文件
- `example_input.xlsx`：包含以下列的 Excel 文件  
  `工作表名称, 编号, 函证单位, 工程项目, 应收帐款（已开票末付款）, 长期应收款（质量保金）, 合计`
- `template.xlsx`：询证函模板（程序会自动读取初始值）

### 2. 运行程序
- **Windows**：双击 `inquiry-letter-generator.exe`
- **MacOS**：空了来添加【to-do】

### 3. 操作流程
1. 点击 **“选择询证函台账文件”**，选择你的 Excel 台账
2. 编辑模板字段（回函地址、联系人、电话、邮箱、发函日期）
3. 点击 **“生成询证函.xlsx”**，选择保存位置
4. 程序将生成带格式的询证函，并自动更新本地模板

## ⚙️ 开发与打包

### 依赖
```txt
PyQt6
pandas
openpyxl
```

### 本地运行
```
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
python extract_excel_qt.py
```

### 自动打包

本项目使用 GitHub Actions 自动打包：

推送代码到 main 分支 → 自动构建 Windows .exe
工件下载地址：Actions > Build Windows Executable

### 工作原理

程序会将台账中的数据应用到固定格式的询证函中，询证函的模版与台账对应的单元格如下：
  |询证函模板中的单元格|台账中的列|
  |---|---|
  |D1|编号|
  |A3|函证单位|
  |A4(部分)|项目|
  |C13|应收帐款（已开票末付款）|
  |C14'|长期应收款（质量保金）|
  |C16|合计|

模版中可编辑的数据：
  |询证函模板中的单元格|数据|
  |---|---|
  |A9|回函地址、联系人|
  |A10|电话|
  |C10|邮箱|
  |B19|发函单位|
  |D20|日期|

## 📎 示例文件

- [`template.xlsx`](template.xlsx)：标准询证函模板（含合并单元格、格式）
- [`example_input.xlsx`](example_input.xlsx)：示例台账文件（包含所有必要列）

## 📝 作者

KevinMai