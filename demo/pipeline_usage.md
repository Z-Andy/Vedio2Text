# 三段式视频流程使用说明

## API Key 分工

| 文件 | 用途 | 默认模型 |
| --- | --- | --- |
| `../apikey1.txt` | 生成视频 | `sora-2` + `1280x720`，网关实际命中 `sora-2_1280x720` |
| `../apikey2.txt` | 视频三层事实描述 | `doubao-seed-2-0-lite-260428` |
| `../apikey3.txt` | 裁判打标 | `deepseek-v4-flash` |

`apikey*.txt` 支持两种格式：

```json
{"key":"sk-...","url":"https://..."}
```

或：

```json
"key":"sk-...","url":"https://..."
```

## 单独生成视频

```powershell
powershell -ExecutionPolicy Bypass -File .\generate_video.ps1 -PromptPath .\Prompt1.md -OutputPath .\Vedio1.mp4
```

## 单独分析视频

```powershell
powershell -ExecutionPolicy Bypass -File .\analyze_video.ps1 -VideoPath .\Vedio1.mp4 -OriginalPromptPath .\Prompt1.md
```

输出目录：

```text
analysis/Vedio1/
```

其中包含：

- `macro.json`
- `segment.json`
- `frame.json`

分析脚本会在保存时去掉模型可能附带的 Markdown 代码块标记，输出文件应为可直接解析的合法 JSON。

三层分析会把原始视频生成提示词一起传入，用于提取需要重点观察的核心要素。分析层只记录事实状态，例如 `有`、`无`、`不可见`、`未知`，不在分析阶段判断是否错误。

## 单独裁判打标

```powershell
powershell -ExecutionPolicy Bypass -File .\judge_labels.ps1 -VideoName Vedio1 -OriginalPromptPath .\Prompt1.md -UseSimpleUserPrompt
```

建议使用 `-UseSimpleUserPrompt`。裁判脚本仍会逐一级标签、逐层输入打标，并把已有打标结果传给下一批用于去重、替换、保留或少量并列。三层描述不会被截断，`-MaxTokens` 只限制模型输出长度。

最终成果写入 `labels/<视频名>/final_labels.json`，格式为数组，每项只保留：

- `一级标签--二级标签`
- `证据来源`
- `错因描述`

只测试一个一级标签和一个层级：

```powershell
powershell -ExecutionPolicy Bypass -File .\judge_labels.ps1 `
  -VideoName Vedio1 `
  -OriginalPromptPath .\Prompt1.md `
  -Categories '提示词一致性' `
  -LayerNames '宏观层' `
  -MaxTokens 800 `
  -UseSimpleUserPrompt
```

输出目录：

```text
labels/Vedio1/
```

最终标签文件：

```text
labels/Vedio1/final_labels.json
```

## 批量运行三条视频

```powershell
powershell -ExecutionPolicy Bypass -File .\run_pipeline.ps1 -SkipExistingAnalysis -UseSimpleJudgePrompt
```

默认处理：

- `Vedio1.mp4` + `Prompt1.md`
- `Vedio2.mp4` + `Prompt2.md`
- `Vedio3.mp4` + `Prompt3.md`

## 流水线并行运行

```powershell
powershell -ExecutionPolicy Bypass -File .\run_pipeline_parallel.ps1 `
  -SkipExistingAnalysis `
  -UseSimpleJudgePrompt `
  -MaxConcurrentAnalysis 3 `
  -MaxConcurrentJudge 1 `
  -JudgeMaxTokens 4000
```

并行脚本会让不同视频的分析和裁判流水线运作。单个视频内部的裁判仍按顺序执行，保证下一批标签能看到上一批已有打标结果。

## DeepSeek 连通性说明

当前裁判脚本默认开启 `deepseek-v4-flash` 思考模式，并要求返回 JSON 对象。因为思考模式会消耗更多输出额度，建议 `JudgeMaxTokens` 使用 `4000` 或更高。

如需调试速度或避免思考内容占用输出额度，可以额外传入：

```powershell
-DisableJudgeThinking
```
