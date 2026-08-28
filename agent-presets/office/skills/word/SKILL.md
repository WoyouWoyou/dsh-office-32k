---
name: word
description: 处理 Word 文档（.docx）：读取与提取（段落、表格、标题结构）、修改编辑（替换文本、改样式、增删段落与表格）、新建生成正式文档。全部依赖来自离线包 D:\offline-deps-example，禁止联网安装。
---

# Word 处理技能

**离线依赖原则**：依赖一律来自 `D:\offline-deps-example`。
禁止 `pip install`（联网）、禁止下载、禁止 Word/Excel COM、禁止 tempfile/mkdtemp。

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
if (-not (Test-Path (Join-Path $deps 'docx'))) {
  New-Item -ItemType Directory -Path $deps -Force | Out-Null
  Get-ChildItem $wh -Filter 'python_docx-*.whl' | ForEach-Object { python -m zipfile -e $_.FullName $deps }
  Get-ChildItem $wh -Filter 'openpyxl-*.whl' | ForEach-Object { python -m zipfile -e $_.FullName $deps }
  Get-ChildItem $wh -Filter 'et_xmlfile-*.whl' | ForEach-Object { python -m zipfile -e $_.FullName $deps }
  Get-ChildItem $wh -Filter 'typing_extensions-*.whl' | ForEach-Object { python -m zipfile -e $_.FullName $deps }
}
# ③ 需要 pandas/numpy 时才切换所选根目录的便携 Python 3.11（已装）；否则用系统 python
$env:OFFICE_PY = 'python'
if ((Test-Path $offlinePy) -and (& $offlinePy -c "import pandas" 2>$null; $LASTEXITCODE -eq 0)) {
  $env:OFFICE_PY = $offlinePy
}
& $env:OFFICE_PY -c "import docx; print('python-docx OK')"
```

**环境就绪后立即开始任务**。禁止：继续装包/解压/验证循环、解压文件名含
`cp311`/`cp312`/`cp313` 的编译 wheel、把文件写到工作区之外、pip、联网、COM、tempfile。
之后一律用 `& $env:OFFICE_PY 脚本.py` 运行（脚本写成文件再跑，不要 `-c` 贴大段代码）。
脚本开头加一行 `import sys; sys.stdout.reconfigure(encoding='utf-8', errors='replace')`
避免中文文件名在控制台乱码。
老版 .doc（二进制）：不要用 COM，请用户先另存为 .docx。

## 1. 读取与提取（小步：先看结构）

**先看文档结构（标题大纲 + 段落/表格数），不要全文打印**：
```python
from docx import Document
doc = Document(r'文件.docx')
print('paragraphs:', len(doc.paragraphs), 'tables:', len(doc.tables))
for p in doc.paragraphs:
    if p.style.name.startswith(('Heading', '标题')):
        print(p.style.name, '|', p.text[:60])
```
**按需提取**：正文段落循环打印（可加筛选）；表格 `tbl.rows`/`tbl.columns` 取单元格文本。
**大文档**：先输出大纲，用户确认后再提取具体章节；整篇内容写 .txt 文件，不贴对话。

## 2. 修改与编辑

**绝不直接改原文件**，另存 `_修改.docx`：
```python
from docx import Document
doc = Document(r'文件.docx')
for p in doc.paragraphs:
    if '旧词' in p.text:
        for run in p.runs:
            run.text = run.text.replace('旧词', '新词')
for tbl in doc.tables:
    for row in tbl.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.text = run.text.replace('旧词', '新词')
doc.save(r'文件_修改.docx')
```
追加段落 `doc.add_paragraph('...')`；删除表格用 `tbl._tbl.getparent().remove(tbl._tbl)`。

## 3. 新建与生成（中文必须显式设字体）

```python
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

doc = Document()
style = doc.styles['Normal']
style.font.name = 'Microsoft YaHei'
style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
doc.add_heading('报告标题', level=0)
doc.add_paragraph('正文内容……')
tbl = doc.add_table(rows=2, cols=3)
tbl.style = 'Table Grid'
tbl.rows[0].cells[0].text = '列1'
doc.save(r'输出报告.docx')
```

## 4. 细粒度执行纪律（32K 上下文关键）

- **先小后大**：先写最小脚本生成 1 段/1 表的样例验证，再生成完整文档；报错改小脚本重跑。
- **一个脚本只干一件事**；脚本、中间文本、最终 .docx 都落盘，断点续跑。
- 每次命令输出控制在几十行内；长内容写文件。
- 完成汇报：输入文件、改动内容、输出文件路径、关键摘要。
