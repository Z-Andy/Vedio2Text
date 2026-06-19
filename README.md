# Vedio2Text

一个面向 AI 生成视频质量评估的分层证据驱动流水线。项目先把视频内容转写为宏观层、片段层、逐帧层三类事实描述，再基于原始生成提示词和标签体系进行裁判打标，最终输出可追溯的结构化 JSON 标签结果。

> 项目目录名和脚本中沿用了 `Vedio` 拼写，运行命令请按现有文件名使用。

## 核心思路

本项目采用“先描述，后裁判”的两阶段流程：

1. 视频理解层：使用多模态模型读取视频和原始提示词，生成三层事实描述。
2. 标签裁判层：使用裁判模型结合原始提示词、三层描述和标签体系，判断视频是否命中质量问题标签。

这种方式避免直接让模型端到端给结论，便于人工复核每个标签对应的证据来源和错因描述。

## 目录结构

```text
.
+-- demo/                       # PowerShell 流水线脚本
|   +-- generate_video.ps1       # 根据提示词生成视频
|   +-- analyze_video.ps1        # 生成三层视频事实描述
|   +-- judge_labels.ps1         # 基于三层描述进行标签裁判
|   +-- run_pipeline.ps1         # 顺序批量运行分析和裁判
|   +-- run_pipeline_parallel.ps1# 并行批量运行分析和裁判
|   +-- pipeline_usage.md        # 原始使用说明
+-- Prompt/                      # 提示词、标签体系和三层权责文档
|   +-- macro_prompt.md
|   +-- segment_prompt.md
|   +-- frame_prompt.md
|   +-- judge_prompt.md
|   +-- judge_prompt_compact.md
|   +-- labels.json
|   +-- 三层权责文档.md
+-- sample/                      # 示例视频、提示词、分析结果和标签结果
|   +-- Vedio.mp4
|   +-- Prompt.md
|   +-- analysis/
|   +-- labels/
|   +-- label_metrics_dashboard.html
+-- tools/                       # 汇报材料生成辅助脚本
+-- 汇报PPT/                     # 项目汇报材料
+-- 项目记录.md                  # 项目过程记录
```

## 环境要求

- Windows PowerShell
- 可访问对应模型服务的 API Key
- 网络环境能连接配置的模型服务地址

脚本不依赖额外的 Python 或 Node.js 包，主要通过 PowerShell 调用 HTTP API。

## API Key 配置

项目根目录下使用三个 API Key 文件，已在 `.gitignore` 中忽略：

| 文件 | 用途 | 默认模型 |
| --- | --- | --- |
| `apikey1.txt` | 视频生成 | `sora-2` |
| `apikey2.txt` | 视频三层分析 | `doubao-seed-2-0-lite-260428` |
| `apikey3.txt` | 裁判打标 | `deepseek-v4-flash` |

推荐写成 JSON：

```json
{"key":"sk-xxx","url":"https://api.example.com/v1"}
```

也可以写成脚本支持的简写形式：

```text
"key":"sk-xxx","url":"https://api.example.com/v1"
```

部分脚本也支持只写 API Key，此时会使用脚本中的默认服务地址。

## 快速开始

进入脚本目录：

```powershell
cd .\demo
```

### 1. 单独生成视频

```powershell
powershell -ExecutionPolicy Bypass -File .\generate_video.ps1 `
  -PromptPath .\Prompt1.md `
  -OutputPath .\Vedio1.mp4
```

常用参数：

- `-Model`：视频生成模型，默认 `sora-2`
- `-Size`：视频尺寸，默认 `1280x720`
- `-Seconds`：视频时长，默认 `8`
- `-Proxy`：代理地址，可选

### 2. 单独分析视频

```powershell
powershell -ExecutionPolicy Bypass -File .\analyze_video.ps1 `
  -VideoPath .\Vedio1.mp4 `
  -OriginalPromptPath .\Prompt1.md
```

输出目录：

```text
demo/analysis/Vedio1/
+-- macro.json
+-- segment.json
+-- frame.json
+-- macro_response.json
+-- segment_response.json
+-- frame_response.json
```

其中：

- `macro.json`：宏观层描述，关注全片稳定成立的主体、场景、动作、风格和整体呈现。
- `segment.json`：片段层描述，关注时间轴上的主体延续、动作过程、场景变化和音画关系。
- `frame.json`：逐帧层描述，关注关键帧或高频采样帧中的局部结构、文字、画面质量和可见缺陷。

### 3. 单独裁判打标

```powershell
powershell -ExecutionPolicy Bypass -File .\judge_labels.ps1 `
  -VideoName Vedio1 `
  -OriginalPromptPath .\Prompt1.md `
  -UseSimpleUserPrompt
```

输出目录：

```text
demo/labels/Vedio1/
```

最终标签文件：

```text
demo/labels/Vedio1/final_labels.json
```

最终结果会压缩为数组，每个标签保留：

- `一级标签--二级标签`
- `证据来源`
- `错因描述`

### 4. 顺序批量运行流水线

默认处理：

- `Vedio1.mp4` + `Prompt1.md`
- `Vedio2.mp4` + `Prompt2.md`
- `Vedio3.mp4` + `Prompt3.md`

```powershell
powershell -ExecutionPolicy Bypass -File .\run_pipeline.ps1 `
  -SkipExistingAnalysis `
  -UseSimpleJudgePrompt
```

也可以指定视频名：

```powershell
powershell -ExecutionPolicy Bypass -File .\run_pipeline.ps1 `
  -VideoNames Vedio1,Vedio2 `
  -SkipExistingAnalysis `
  -UseSimpleJudgePrompt
```

### 5. 并行批量运行流水线

```powershell
powershell -ExecutionPolicy Bypass -File .\run_pipeline_parallel.ps1 `
  -SkipExistingAnalysis `
  -UseSimpleJudgePrompt `
  -MaxConcurrentAnalysis 3 `
  -MaxConcurrentJudge 1 `
  -JudgeMaxTokens 4000
```

并行脚本会让不同视频的分析和裁判任务并行执行。单个视频内部的裁判仍按顺序进行，以保证后续批次能看到前面已有的打标结果，用于去重、替换或保留。

## 输出说明

分析阶段输出三类 JSON：

| 文件 | 含义 |
| --- | --- |
| `macro.json` | 全片层面的事实描述 |
| `segment.json` | 按时间片段组织的事实描述 |
| `frame.json` | 关键帧或高频采样帧描述 |

裁判阶段输出：

| 文件 | 含义 |
| --- | --- |
| `*_response.json` | 模型原始响应 |
| `*.json` | 单批次裁判结果 |
| `final_labels.json` | 汇总后的最终标签 |

示例结果可参考：

```text
sample/analysis/
sample/labels/final_labels.json
sample/label_metrics_dashboard.html
```

## 标签体系

标签体系位于：

```text
Prompt/labels.json
```

当前覆盖的一级维度包括：

- 提示词一致性
- 视觉生成质量
- 时间一致性
- 物理真实性
- 音频生成质量
- 审美质量
- 安全与合规

如需迁移到其他视频评估任务，通常优先修改：

- `Prompt/labels.json`
- `Prompt/macro_prompt.md`
- `Prompt/segment_prompt.md`
- `Prompt/frame_prompt.md`
- `Prompt/judge_prompt_compact.md`

## 常用参数

`analyze_video.ps1`：

| 参数 | 说明 |
| --- | --- |
| `-ApiKeyPath` | 分析模型 API Key 文件，默认 `..\apikey2.txt` |
| `-VideoPath` | 待分析视频路径 |
| `-OriginalPromptPath` | 原始视频生成提示词 |
| `-OutputDir` | 分析结果输出目录，默认 `.\analysis` |
| `-PromptDir` | 三层分析提示词目录，默认 `..\Prompt` |
| `-Model` | 分析模型 |
| `-Proxy` | 代理地址 |

`judge_labels.ps1`：

| 参数 | 说明 |
| --- | --- |
| `-ApiKeyPath` | 裁判模型 API Key 文件，默认 `..\apikey3.txt` |
| `-VideoName` | 视频名称，不含扩展名 |
| `-OriginalPromptPath` | 原始视频生成提示词 |
| `-AnalysisDir` | 三层分析结果目录，默认 `.\analysis` |
| `-OutputDir` | 标签输出目录，默认 `.\labels` |
| `-LayerMode` | `Separate` 或 `Combined` |
| `-Categories` | 只评估指定一级标签 |
| `-LayerNames` | 只评估指定层级 |
| `-MaxTokens` | 裁判模型最大输出长度 |
| `-UseSimpleUserPrompt` | 使用更短的用户提示词 |
| `-DisableThinking` | 关闭支持该参数的模型思考模式 |

## 注意事项

- 运行 `demo` 下的脚本时，建议先 `cd demo`，因为脚本默认用相对路径寻找 `..\Prompt` 和 `..\apikey*.txt`。
- `apikey*.txt` 不应提交到版本库。
- 分析脚本会把视频转为 Base64 data URL 后上传，视频过大时可能导致请求体过大或接口超时。
- 裁判阶段建议使用 `-UseSimpleUserPrompt` 和较高的 `-MaxTokens`，例如 `4000`。
- 若只想复用已有分析结果，可给批量脚本传入 `-SkipExistingAnalysis`。

## 参考材料

- [demo/pipeline_usage.md](demo/pipeline_usage.md)
- [项目记录.md](项目记录.md)
- [Prompt/三层权责文档.md](Prompt/三层权责文档.md)
