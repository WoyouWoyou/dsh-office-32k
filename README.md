# DSH · 办公文档·32K 模式（Agent Preset 插件）

为 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) 打造的 **32K 上下文办公文档 Agent 预设**：
在窄上下文（32768 = 输入+输出，输出上限 8192）的本地模型上，稳定地处理 **Excel / Word / PPT / CSV**，
全程离线、数据不出本机。

## 特性

- **32K 上下文管理**：针对"中文 token 被估算器低估约 2 倍 + 8K 输出预算"做了压缩策略校正
  （阈值 0.3 → 估算约 9.8K 触发，保留 5K，工具结果裁剪 8K 字符），配合溢出恢复兜底。
- **离线依赖**：所有 Python 依赖来自本地离线 wheelhouse（`OFFICE_DEPS_ROOT` 指定）；
  纯 Python wheel 直接解压到工作区 `.office-deps` 即可用，编译包走便携 Python 3.11。
- **办公技能**：excel（openpyxl/pandas，含易错速查）、word（python-docx），
  另支持 python-pptx / docxtpl / matplotlib / pypdf 等（离线 wheelhouse 已收录）。
- **细粒度任务分解**：人设强制"探查 → 最小脚本 → 小样验证 → 全量 → 校验 → 汇报"，
  每步落盘、断点续跑、只贴关键错误行。
- **工具精简**：移除对本地 32B 模型易出错的 `str_replace_editor`（其参数键名常被记错导致
  反复 INVALID_ARGS），脚本工作统一走 `write`/`edit`。

## 目录结构

```
agent-presets/office/          # 插件本体（DSH Agent Preset）
├── agent.cordis.yml           # 组装：人设 + 工具 + 32K 压缩策略 + 技能发现
├── preset.yml                 # 显示名「办公文档·32K 模式」
└── skills/
    ├── excel/SKILL.md         # Excel 技能
    └── word/SKILL.md          # Word 技能
scripts/
├── repack-portable.ps1        # 便携版重新打包脚本（剔除运行时产物、零链接校验、tar 打包）
└── analyze-session.mjs        # 会话日志工具调用统计分析
```

## 安装

把 `agent-presets/office` 目录放入 DSH 的预设目录：

- **源码版**：`<checkout>/apps/cli/config/agent-presets/office`
- **便携版**：`<portable>/dsh/config/agent-presets/office`

预设发现是实时读盘的，**无需重启**，刷新浏览器即可在新建会话的预设列表看到「办公文档·32K 模式」。

## 使用

1. 新建会话 → 预设选「办公文档·32K 模式」；或在 `settings.yaml` 设默认：
   ```yaml
   agent-presets:
     default: office
   ```
2. 模型配置（你的模型真实窗口/输出上限）：
   ```yaml
   llm-pi-ai:
     providers:
       your-model:
         api: openai-completions
         baseURL: http://your-model-gateway/v1
         models:
           - id: your-model-id
             contextWindow: 32768
             maxTokens: 8192
   ```
3. 离线依赖：设置环境变量 `OFFICE_DEPS_ROOT` 指向含 `deps/wheelhouse` 的离线包根目录；
   不设置时回退到示例路径 `D:\offline-deps-example`（请改成你的实际路径）。

## 离线依赖说明

- **纯 Python wheel**（docx/openpyxl/pptx/docxtpl/xlsxwriter/xlrd/jieba/tabulate/pypdf/bs4 等）：
  直接 `python -m zipfile -e <wheel> <工作区>/.office-deps` 解压 + `PYTHONPATH` 即可，任何 Python 可用。
- **编译 wheel**（pandas/numpy/matplotlib/scipy，cp311）：需 Python 3.11（便携运行时），
  技能会优先选择已装这些包的便携 Python。
- 依赖根目录可随时切换（改 `OFFICE_DEPS_ROOT`），技能引导已内置该逻辑。

## 实测对比（同一 32B 模型、同一"修复脚本"任务）

| 指标 | 改动前（有 str_replace_editor） | 改动后（已移除） |
|---|---|---|
| 工具调用次数 | 13 | **6（−54%）** |
| 执行步数 | 13 | **6（−54%）** |
| 相邻重试 | 7 | **2（−71%）** |
| 模型输出量 | 137 | **66（−52%）** |

## 许可

[MIT](LICENSE)
