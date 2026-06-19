param(
  [string[]]$VideoNames = @("Vedio1", "Vedio2", "Vedio3"),
  [string]$AnalysisModel = "doubao-seed-2-0-lite-260428",
  [string]$JudgeModel = "deepseek-v4-flash",
  [ValidateSet("Separate", "Combined")]
  [string]$LayerMode = "Separate",
  [int]$JudgeRequestTimeoutSec = 240,
  [int]$JudgeMaxTokens = 4000,
  [int]$JudgeMaxRetries = 1,
  [int]$MaxConcurrentAnalysis = 3,
  [int]$MaxConcurrentJudge = 2,
  [switch]$UseSimpleJudgePrompt,
  [switch]$DisableJudgeThinking,
  [switch]$SkipExistingAnalysis,
  [switch]$SkipExistingLabels
)

$ErrorActionPreference = "Stop"

function Get-PromptPathForVideo {
  param([string]$VideoName)

  $index = [regex]::Match($VideoName, "\d+").Value
  if ([string]::IsNullOrWhiteSpace($index)) {
    throw "Cannot infer prompt index from video name: $VideoName"
  }

  return ".\Prompt$index.md"
}

function Test-AnalysisComplete {
  param([string]$VideoName)

  $analysisPath = Join-Path ".\analysis" $VideoName
  $required = @("macro.json", "segment.json", "frame.json") | ForEach-Object { Join-Path $analysisPath $_ }
  return (($required | Where-Object { -not (Test-Path -LiteralPath $_) }).Count -eq 0)
}

function Test-LabelsComplete {
  param([string]$VideoName)

  return (Test-Path -LiteralPath (Join-Path ".\labels\$VideoName" "final_labels.json"))
}

function Start-AnalysisJobForVideo {
  param([string]$VideoName)

  $videoPath = ".\$VideoName.mp4"
  $promptPath = Get-PromptPathForVideo -VideoName $VideoName
  Write-Host "Start analysis job: $VideoName"
  return Start-Job -Name "analysis_$VideoName" -ScriptBlock {
    param($WorkDir, $VideoPath, $PromptPath, $AnalysisModel)
    Set-Location $WorkDir
    powershell -ExecutionPolicy Bypass -File .\analyze_video.ps1 -VideoPath $VideoPath -OriginalPromptPath $PromptPath -Model $AnalysisModel
    exit $LASTEXITCODE
  } -ArgumentList (Get-Location).Path, $videoPath, $promptPath, $AnalysisModel
}

function Start-JudgeJobForVideo {
  param([string]$VideoName)

  $promptPath = Get-PromptPathForVideo -VideoName $VideoName
  Write-Host "Start judge job: $VideoName"
  return Start-Job -Name "judge_$VideoName" -ScriptBlock {
    param($WorkDir, $VideoName, $PromptPath, $JudgeModel, $LayerMode, $JudgeRequestTimeoutSec, $JudgeMaxTokens, $JudgeMaxRetries, $UseSimpleJudgePrompt, $DisableJudgeThinking)
    Set-Location $WorkDir
    $args = @(
      "-ExecutionPolicy", "Bypass",
      "-File", ".\judge_labels.ps1",
      "-VideoName", $VideoName,
      "-OriginalPromptPath", $PromptPath,
      "-Model", $JudgeModel,
      "-LayerMode", $LayerMode,
      "-RequestTimeoutSec", $JudgeRequestTimeoutSec,
      "-MaxTokens", $JudgeMaxTokens,
      "-MaxRetries", $JudgeMaxRetries,
      "-JudgePromptFile", "judge_prompt_compact.md",
      "-ContinueOnBatchError"
    )
    if ([bool]$UseSimpleJudgePrompt) {
      $args += "-UseSimpleUserPrompt"
    }
    if ([bool]$DisableJudgeThinking) {
      $args += "-DisableThinking"
    }
    powershell @args
    exit $LASTEXITCODE
  } -ArgumentList (Get-Location).Path, $VideoName, $promptPath, $JudgeModel, $LayerMode, $JudgeRequestTimeoutSec, $JudgeMaxTokens, $JudgeMaxRetries, $UseSimpleJudgePrompt, $DisableJudgeThinking
}

$pendingAnalysis = New-Object System.Collections.Queue
foreach ($name in $VideoNames) {
  if ($SkipExistingAnalysis -and (Test-AnalysisComplete -VideoName $name)) {
    Write-Host "Analysis already complete: $name"
    continue
  }
  $pendingAnalysis.Enqueue($name)
}

$pendingJudge = New-Object System.Collections.Queue
foreach ($name in $VideoNames) {
  if ((Test-AnalysisComplete -VideoName $name) -and -not ($SkipExistingLabels -and (Test-LabelsComplete -VideoName $name))) {
    $pendingJudge.Enqueue($name)
  }
}

$analysisJobs = @{}
$judgeJobs = @{}
$completedAnalysis = @{}
$completedJudge = @{}

while ($pendingAnalysis.Count -gt 0 -or $analysisJobs.Count -gt 0 -or $pendingJudge.Count -gt 0 -or $judgeJobs.Count -gt 0) {
  while ($pendingAnalysis.Count -gt 0 -and $analysisJobs.Count -lt $MaxConcurrentAnalysis) {
    $name = $pendingAnalysis.Dequeue()
    $job = Start-AnalysisJobForVideo -VideoName $name
    $analysisJobs[$job.Id] = @{
      Name = $name
      Job = $job
    }
  }

  foreach ($id in @($analysisJobs.Keys)) {
    $entry = $analysisJobs[$id]
    $job = $entry.Job
    if ($job.State -in @("Completed", "Failed", "Stopped")) {
      Receive-Job -Job $job
      $exitCode = $job.ChildJobs[0].JobStateInfo.Reason
      Remove-Job -Job $job
      $analysisJobs.Remove($id)

      if (Test-AnalysisComplete -VideoName $entry.Name) {
        Write-Host "Analysis complete: $($entry.Name)"
        $completedAnalysis[$entry.Name] = $true
        if (-not ($SkipExistingLabels -and (Test-LabelsComplete -VideoName $entry.Name))) {
          $pendingJudge.Enqueue($entry.Name)
        }
      }
      else {
        Write-Host "Analysis failed or incomplete: $($entry.Name)"
      }
    }
  }

  while ($pendingJudge.Count -gt 0 -and $judgeJobs.Count -lt $MaxConcurrentJudge) {
    $name = $pendingJudge.Dequeue()
    if ($completedJudge.ContainsKey($name)) {
      continue
    }
    $job = Start-JudgeJobForVideo -VideoName $name
    $judgeJobs[$job.Id] = @{
      Name = $name
      Job = $job
    }
  }

  foreach ($id in @($judgeJobs.Keys)) {
    $entry = $judgeJobs[$id]
    $job = $entry.Job
    if ($job.State -in @("Completed", "Failed", "Stopped")) {
      Receive-Job -Job $job
      Remove-Job -Job $job
      $judgeJobs.Remove($id)

      if (Test-LabelsComplete -VideoName $entry.Name) {
        Write-Host "Judging complete: $($entry.Name)"
        $completedJudge[$entry.Name] = $true
      }
      else {
        Write-Host "Judging failed or incomplete: $($entry.Name)"
      }
    }
  }

  Start-Sleep -Seconds 2
}

Write-Host "Parallel pipeline finished."
