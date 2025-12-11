# PDF-Data-Recognition-and-Organization

一个基于 PySide6 的桌面工具，用于 PDF 与 Word 互转，以及从医学报告 PDF 中结构化提取数据并导出为 Excel。界面采用现代深色主题，支持拖拽文件、默认文件名、进度提示与日志输出。

## 功能概览

- 文件转换（合并栏目）
  - PDF 转 Word（`pdf2docx`）
  - Word 转 PDF（`docx2pdf`，Windows 需安装 Microsoft Word）
- 报告提取（合并栏目）
  - 基础信息：姓名、性别、年龄、采样日期、标本类型、住院号、病理号、身份证号、送检日期、病历号、送检医生、送检单位、检测项目、检测方法、送检材料、临床诊断
  - 重排结果：姓名、检测号、重排基因、左断裂点位、右断裂点位（优先使用 `pdfplumber` 表格提取，配合稳健回退）
  - 突变数据：检测号、突变基因、转录本 ID、外显子、核苷酸改变、氨基酸改变、突变频率（`pdfplumber` 动态列匹配）

## 界面亮点

- 两大主栏目：
  - “文件转换”下拉选择 PDF→Word / Word→PDF
  - “报告提取”下拉选择 基础信息 / 重排结果 / 突变数据
- 深色主题：统一的卡片式布局、主次按钮、圆角与层次阴影（见 `theme.py`）
- 拖拽文件：在卡片区域拖拽文件或点击“浏览…”选择
- 默认保存名：基础信息.xlsx、重排结果.xlsx、突变数据.xlsx
- 线程执行：转换与解析在后台线程运行，避免界面卡顿
- 日志面板：显示已选文件、成功与错误信息，便于排查

## 目录结构

- `mainwindow.py`：应用入口与主窗口，加载主题与两个合并栏目
- `ui_components.py`：界面组件（拖拽卡片、日志面板、合并栏目 `CombinedConversionTab`、单功能视图 `ConversionTab`）
- `workers.py`：后台任务
  - PDF→Word（`pdf2docx`）
  - Word→PDF（`docx2pdf`）
  - 基础信息提取（`PyMuPDF` 文本 + 正则）
  - 重排结果提取（`pdfplumber` 表格 + 多策略回退）
  - 突变数据提取（`pdfplumber` 表格，动态列识别）
- `theme.py`：应用级样式表（QSS）
- `requirements.txt`：依赖清单

## 环境要求

- Python 3.9+（建议）
- Windows（推荐，Word→PDF 需 Microsoft Word）
- 依赖：`PySide6`、`pdf2docx`、`docx2pdf`、`pandas`、`openpyxl`、`PyMuPDF`、`pdfplumber`

## 安装

```bash
pip install -r requirements.txt
```

> 提示：Word→PDF 在 Windows 上需要已安装 Microsoft Word。

## 运行

```bash
python mainwindow.py
```

## 使用说明

- 打开应用后，在顶部标签页选择“文件转换”或“报告提取”。
- 通过下拉选择具体子功能（例如“基础信息”“重排结果”“突变数据”）。
- 将 PDF/Word 文件拖拽到卡片区域或点击“浏览…”选择文件。
- 点击主按钮（如“基础信息”“重排结果”“突变数据”）开始处理，选择保存位置（已预填默认文件名）。
- 等待进度完成，查看提示与日志输出。

## 提取策略说明

- 基础信息：`PyMuPDF` 提取整页文本，正则兼容中文标签间空格（如“姓 名”“性 别”）、同行多字段与变体布局。
- 重排结果：优先 `pdfplumber` 表格解析，依据表头关键字（重排基因/断裂点），动态定位列并清洗占位符（如 “-” → “无”）；必要时回退到文本正则或坐标模式（`chr14:106032614`）。
- 突变数据：`pdfplumber` 动态列匹配（基因/转录本/外显子/c./p./VAF），适配不同中文/英文表头。

## 常见问题

- Word→PDF 报错：请确认系统已安装 Microsoft Word。
- 提取结果为空：请查看日志面板输出，确认 PDF 是否为可复制文本（扫描件可能需 OCR）。
- 表格列未识别：请确认表头包含关键字（如“基因”“改变”“频率”），或者提供样例以优化识别规则。

## 许可证

MIT License

## 致谢

- `PySide6` 提供跨平台 GUI 能力
- `pdf2docx`、`docx2pdf` 提供文档互转能力
- `PyMuPDF`、`pdfplumber` 提供 PDF 文本与表格解析能力
