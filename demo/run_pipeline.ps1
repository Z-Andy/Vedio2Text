param(
  [string[]]$VideoNames = @("Vedio1", "Vedio2", "Vedio3"),
  [string]$AnalysisModel = "doubao-seed-2-0-lite-260428",
  [string]$JudgeModel = "deepseek-v4-flash",
  [switch]$SkipExistingAnalysis,
  [ValidateSet("Separate", "Combined")]
  [string]$LayerMode = "Separate",
  [int]$JudgeRequestTimeoutSec = 240,
  [int]$JudgeMaxTokens = 4000,
  [int]$JudgeMaxRetries = 1,
  [switch]$UseSimpleJudgePrompt,
  [switch]$DisableJudgeThinking
)

$ErrorActionPreference = "Continue"

foreach ($name in $VideoNames) {
  $index = [regex]::Match($name, "\d+").Value
  if ([string]::IsNullOrWhiteSpace($index)) {
    Write-Host "Cannot infer prompt index from video name: $name"
    continue
  }

  $videoPath = ".\$name.mp4"
  $promptPath = ".\Prompt$index.md"

  $analysisPath = Join-Path ".\analysis" $name
  $required = @("macro.json", "segment.json", "frame.json") | ForEach-Object { Join-Path $analysisPath $_ }
  $missing = $required | Where-Object { -not (Test-Path -LiteralPath $_) }

  if ($SkipExistingAnalysis -and $missing.Count -eq 0) {
    Write-Host "==== Analyze $name ===="
    Write-Host "Existing analysis found. Skip analysis."
  }
  else {
    Write-Host "==== Analyze $name ===="
    powershell -ExecutionPolicy Bypass -File .\analyze_video.ps1 `
      -VideoPath $videoPath `
      -OriginalPromptPath $promptPath `
      -Model $AnalysisModel

    if ($LASTEXITCODE -ne 0) {
      Write-Host "Analysis failed for $name. Skip judging because analysis files were not generated."
      continue
    }

    $missing = $required | Where-Object { -not (Test-Path -LiteralPath $_) }
  }

  if ($missing.Count -gt 0) {
    Write-Host "Analysis output is incomplete for $name. Missing:"
    $missing | ForEach-Object { Write-Host "  $_" }
    Write-Host "Skip judging $name."
    continue
  }

  Write-Host "==== Judge $name ===="
  $judgeArgs = @(
    "-ExecutionPolicy", "Bypass",
    "-File", ".\judge_labels.ps1",
    "-VideoName", $name,
    "-OriginalPromptPath", $promptPath,
    "-Model", $JudgeModel,
    "-LayerMode", $LayerMode,
    "-RequestTimeoutSec", $JudgeRequestTimeoutSec,
    "-MaxTokens", $JudgeMaxTokens,
    "-MaxRetries", $JudgeMaxRetries,
    "-JudgePromptFile", "judge_prompt_compact.md",
    "-ContinueOnBatchError"
  )
  if ($UseSimpleJudgePrompt) {
    $judgeArgs += "-UseSimpleUserPrompt"
  }
  if ($DisableJudgeThinking) {
    $judgeArgs += "-DisableThinking"
  }
  powershell @judgeArgs

  if ($LASTEXITCODE -ne 0) {
    Write-Host "Judging failed for $name."
  }
}
