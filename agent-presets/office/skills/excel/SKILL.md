---
name: excel
description: 处理 Excel 电子表格（.xlsx / .xls / .csv）：读取与分析（探查结构、筛选、统计）、修改编辑（单元格、行列、公式、格式）、新建生成报表。全部依赖来自离线包 D:\offline-deps-example，禁止联网安装。
---

# Excel 处理技能

**离线依赖原则**：依赖一律来自 `D:\offline-deps-example`。
禁止 `pip install`（联网）、禁止下载、禁止 COM、禁止 tempfile/mkdtemp。

## 0. 环境准备（工作区已预置 .office-deps，5 秒完成）

```powershell
# 依赖根目录：优先用环境变量 OFFICE_DEPS_ROOT（启动器/配置文件指定），否则用默认离线包
$offlineRoot = $env:OFFICE_DEPS_ROOT
if (-not $offlineRoot) { $offlineRoot = 'D:\offline-deps-example' }
$wh = Join-Path $offlineRoot 'deps\wheelhouse'
$offlinePy = Join-Path $offlineRoot 'deps\runtime\python\python.exe'

# ① 指向工作区已预置的纯 Python 依赖（docx/openpyxl/et_xmlfile/typing_extensions）
$deps = Join-Path (Get-Location) '.office-deps'
$env:PYTHONPATH = "$deps;$env:PYTHONPATH"
# ② 兜底：若换工作目录缺包，从所选依赖根目录的 wheelhouse 解压纯 Python wheel（禁止编译 wheel）
if (-not (Test-Path (Join-Path $deps 'openpyxl'))) {
  New-Item -ItemType Directory -Path $deps -Force | Out-Null
  Get-ChildItem $wh -Filter 'openpyxl-*.whl' | ForEach-Object { python -m zipfile -e $_.FullName $deps }
  Get-ChildItem $wh -Filter 'et_xmlfile-*.whl' | ForEach-Object { python -m zipfile -e $_.FullName $deps }
  Get-ChildItem $wh -Filter 'python_docx-*.whl' | ForEach-Object { python -m zipfile -e $_.FullName $deps }
  Get-ChildItem $wh -Filter 'typing_extensions-*.whl' | ForEach-Object { python -m zipfile -e $_.FullName $deps }
}
# ③ 需要 pandas/numpy 时才切换所选根目录的便携 Python 3.11（已装）；否则用系统 python
$env:OFFICE_PY = 'python'
if ((Test-Path $offlinePy) -and (& $offlinePy -c "import pandas" 2>$null; $LASTEXITCODE -eq 0)) {
  $env:OFFICE_PY = $offlinePy
}
& $env:OFFICE_PY -c "import openpyxl; print('openpyxl', openpyxl.__version__)"
```

**环境就绪后立即开始任务**。禁止：继续装包/解压/验证循环、解压文件名含
`cp311`/`cp312`/`cp313` 的编译 wheel、把文件写到工作区之外、pip、联网、COM、tempfile。
之后一律用 `& $env:OFFICE_PY 脚本.py` 运行（脚本写成文件再跑，不要 `-c` 贴大段代码）。

## 1. 读取与分析（小步：先探查，再取样）

**探查结构**（永远先做，避免把大文件全量读进 32K 上下文）：
```python
from openpyxl import load_workbook
wb = load_workbook(r'文件.xlsx', read_only=True, data_only=True)
print('sheets:', wb.sheetnames)
for ws in wb.worksheets:
    print(ws.title, ws.max_row, 'x', ws.max_column)
wb.close()
```
**看前几行确定列含义**：`iter_rows(min_row=1, max_row=5, values_only=True)`，只打印前 5~10 行。
**数据分析**：需要统计/筛选时，把数据交给 pandas（仅离线便携 Python）或 openpyxl 逐行聚合，
只打印结果，不打印原始数据。
- .csv 编码：依次尝试 `utf-8` → `gbk` → `gb18030` → `utf-8-sig`。
- 大表（>5 万行）：openpyxl `read_only=True` 流式读，pandas `chunksize=` 分块。
- 整表给模型理解时：只输出 20~50 行样例 + 统计，完整内容写 .csv/.txt 文件。

## 2. 修改与编辑

用 openpyxl（保留原格式）；**绝不直接改原文件**，先复制或另存 `_修改.xlsx`：
```python
from openpyxl import load_workbook
wb = load_workbook(r'文件.xlsx')          # 注意不要 read_only
ws = wb.active
for row in ws.iter_rows(min_row=2):
    if row[0].value == '目标':
        row[2].value = '新值'
wb.save(r'文件_修改.xlsx')
```
合并单元格 `ws.merge_cells('A1:B2')`；字体 `from openpyxl.styles import Font; cell.font = Font(bold=True)`。

## 3. 新建与生成

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
wb = Workbook(); ws = wb.active; ws.title = '汇总'
ws.append(['姓名', '数量', '金额'])
for c in ws[1]:
    c.font = Font(bold=True); c.fill = PatternFill('solid', fgColor='D9E1F2')
for r in 数据: ws.append(r)
for col in range(1, 4):
    ws.column_dimensions[get_column_letter(col)].width = 14
wb.save(r'输出报表.xlsx')
```

## 4. openpyxl 易错速查（写脚本前先看）

- **新建工作簿用 `Workbook()`；读取已有文件用 `load_workbook(路径)`**。绝不要 `load_workbook()` 不带参数（会报 missing filename）。
- 修改已有文件：`load_workbook(路径)`（不要加 `read_only=True`）；只读分析才用 `read_only=True`。
- 列/行索引从 1 开始：`ws.cell(row, col)` 或 `ws['A1']`；`ws.max_row` / `ws.max_column`。
- 读数值注意可能是字符串/公式：必要时 `float()` 转换并捕获异常；汇总前 `total += float(v or 0)`。
- 表头定位用首行文本匹配，不要硬编码列号（列顺序可能变）。

## 5. 细粒度执行纪律（32K 上下文关键）

- **先小后大**：先写只处理前 5 行的最小脚本验证正确，再全量执行；报错时改小脚本重跑。
- **一个脚本只干一件事**；步骤之间产物落盘（脚本、中间 csv、结果 xlsx），断点续跑。
- 每次命令输出控制在几十行内；长输出写文件。
- 完成汇报：输入文件、做了什么、输出文件路径、关键结果摘要。
