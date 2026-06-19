param(
  [string]$ApiKeyPath = "..\apikey3.txt",
  [string]$VideoName = "Vedio1",
  [string]$OriginalPromptPath = ".\Prompt1.md",
  [string]$AnalysisDir = ".\analysis",
  [string]$OutputDir = ".\labels",
  [string]$PromptDir = "..\Prompt",
  [string]$JudgePromptFile = "judge_prompt_compact.md",
  [string]$Model = "deepseek-v4-flash",
  [string]$Proxy = "",
  [string[]]$Categories = @(),
  [string[]]$LayerNames = @(),
  [ValidateSet("Separate", "Combined")]
  [string]$LayerMode = "Separate",
  [int]$RequestTimeoutSec = 240,
  [int]$MaxTokens = 4000,
  [int]$MaxRetries = 1,
  [switch]$UseSimpleUserPrompt,
  [switch]$DisableThinking,
  [switch]$ContinueOnBatchError
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName System.Web.Extensions

function ConvertTo-JsonStringLiteral {
  param([string]$Text)

  $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
  $serializer.MaxJsonLength = 67108864
  return $serializer.Serialize($Text)
}

function ConvertTo-SmallJson {
  param($Value, [int]$Depth = 30)

  if ($null -eq $Value) {
    return "null"
  }

  return (ConvertTo-Json -InputObject $Value -Depth $Depth -Compress)
}

function ConvertTo-LayerInputJson {
  param([System.Collections.IDictionary]$InputObject)

  $parts = New-Object System.Collections.ArrayList
  foreach ($key in $InputObject.Keys) {
    $nameJson = ConvertTo-JsonStringLiteral -Text ([string]$key)
    $value = $InputObject[$key]

    if ($value -is [string]) {
      $valueJson = ConvertTo-JsonStringLiteral -Text $value
    }
    elseif ([string]$key -eq "已有打标结果") {
      $valueJson = ConvertTo-SmallJson -Value $value -Depth 80
    }
    else {
      $valueJson = ConvertTo-SmallJson -Value $value -Depth 30
    }

    [void]$parts.Add("$nameJson`:$valueJson")
  }

  return "{" + ($parts -join ",") + "}"
}

function ConvertTo-ChatBodyJson {
  param(
    [string]$Model,
    [object[]]$Messages,
    [int]$MaxTokens,
    [bool]$DisableThinking
  )

  $messageParts = New-Object System.Collections.ArrayList
  foreach ($message in $Messages) {
    $roleJson = ConvertTo-JsonStringLiteral -Text ([string]$message.role)
    $contentJson = ConvertTo-JsonStringLiteral -Text ([string]$message.content)
    [void]$messageParts.Add("{""role"":$roleJson,""content"":$contentJson}")
  }

  $modelJson = ConvertTo-JsonStringLiteral -Text $Model
  $thinkingType = "enabled"
  if ($DisableThinking) {
    $thinkingType = "disabled"
  }

  return "{""model"":$modelJson,""messages"":[" + ($messageParts -join ",") + "],""temperature"":0,""max_tokens"":$MaxTokens,""thinking"":{""type"":""$thinkingType""},""response_format"":{""type"":""json_object""}}"
}

function Read-ApiConfig {
  param([string]$Path)

  $raw = (Get-Content -Path $Path -Raw -Encoding UTF8).Trim()
  if ([string]::IsNullOrWhiteSpace($raw)) {
    throw "API config is empty: $Path"
  }

  try {
    $json = $raw | ConvertFrom-Json
  }
  catch {
    try {
      $json = ("{" + $raw.Trim().Trim(",") + "}") | ConvertFrom-Json
    }
    catch {
      if ($raw -like "sk-*") {
        return @{
          Key = $raw
          BaseUrl = "https://api.deepseek.com/v1"
        }
      }

      return @{
        Key = $raw
        BaseUrl = "https://api.openai.com/v1"
      }
    }
  }

  return @{
    Key = [string]$json.key
    BaseUrl = ([string]$json.url).TrimEnd("/")
  }
}

function Get-ResponseText {
  param($Response)

  if ($Response.choices -and $Response.choices.Count -gt 0) {
    $content = [string]$Response.choices[0].message.content
    if (-not [string]::IsNullOrWhiteSpace($content)) {
      return $content
    }

    return [string]$Response.choices[0].message.reasoning_content
  }

  if ($Response.output_text) {
    return [string]$Response.output_text
  }

  return (ConvertTo-Json -InputObject $Response -Depth 30)
}

function Get-ChatCompletionsUrl {
  param([string]$BaseUrl)

  $base = $BaseUrl.TrimEnd("/")
  if ($base -match "/(v1|api/v3)$") {
    return "$base/chat/completions"
  }

  return "$base/v1/chat/completions"
}

function Invoke-JsonPostUtf8 {
  param(
    [string]$Uri,
    [string]$ApiKey,
    [byte[]]$BodyBytes,
    [int]$TimeoutSec,
    [hashtable]$RequestOptions
  )

  $request = [System.Net.HttpWebRequest]::Create($Uri)
  $request.Method = "POST"
  $request.ContentType = "application/json; charset=utf-8"
  $request.Accept = "application/json"
  $request.Headers["Authorization"] = "Bearer $ApiKey"
  $request.Timeout = $TimeoutSec * 1000
  $request.ReadWriteTimeout = $TimeoutSec * 1000
  $request.ContentLength = $BodyBytes.Length

  if ($RequestOptions.ContainsKey("Proxy")) {
    $request.Proxy = New-Object System.Net.WebProxy($RequestOptions.Proxy)
  }

  $requestStream = $request.GetRequestStream()
  try {
    $requestStream.Write($BodyBytes, 0, $BodyBytes.Length)
  }
  finally {
    $requestStream.Dispose()
  }

  try {
    $response = $request.GetResponse()
    try {
      $stream = $response.GetResponseStream()
      $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
      $text = $reader.ReadToEnd()
      return ($text | ConvertFrom-Json)
    }
    finally {
      if ($reader) { $reader.Dispose() }
      if ($stream) { $stream.Dispose() }
      $response.Dispose()
    }
  }
  catch [System.Net.WebException] {
    $errorText = $_.Exception.Message
    if ($_.Exception.Response) {
      $errorStream = $_.Exception.Response.GetResponseStream()
      $errorReader = New-Object System.IO.StreamReader($errorStream, [System.Text.Encoding]::UTF8)
      try {
        $errorText = $errorReader.ReadToEnd()
      }
      finally {
        $errorReader.Dispose()
        $errorStream.Dispose()
        $_.Exception.Response.Dispose()
      }
    }

    throw $errorText
  }
}

function ConvertFrom-ModelJson {
  param([string]$Text)

  $clean = $Text.Trim()
  if ($clean -match '```') {
    $clean = $clean -replace '^```json\s*', ''
    $clean = $clean -replace '^```\s*', ''
    $clean = $clean -replace '\s*```$', ''
  }

  return $clean | ConvertFrom-Json
}

function ConvertTo-CompactLabels {
  param($Labels)

  $compact = New-Object System.Collections.ArrayList
  foreach ($label in @($Labels)) {
    $primary = [string]$label."一级标签"
    $secondary = [string]$label."二级标签"
    $pair = [string]$label."一级标签--二级标签"

    if ([string]::IsNullOrWhiteSpace($pair)) {
      if ([string]::IsNullOrWhiteSpace($primary) -and [string]::IsNullOrWhiteSpace($secondary)) {
        continue
      }

      $pair = "$primary--$secondary"
    }

    [void]$compact.Add([ordered]@{
      "一级标签--二级标签" = $pair
      "证据来源" = $label."证据来源"
      "错因描述" = $label."错因描述"
    })
  }

  return $compact
}

function Save-FinalLabels {
  param(
    [string]$TargetDir,
    $Labels
  )

  $compactLabels = ConvertTo-CompactLabels -Labels $Labels
  $jsonItems = New-Object System.Collections.ArrayList
  foreach ($label in @($compactLabels)) {
    [void]$jsonItems.Add((ConvertTo-Json -InputObject $label -Depth 80))
  }

  $json = "[]"
  if ($jsonItems.Count -gt 0) {
    $json = "[" + ($jsonItems -join ",`n") + "]"
  }

  $json | Set-Content -Path (Join-Path $TargetDir "final_labels.json") -Encoding UTF8
}

function Invoke-Judge {
  param(
    [string]$BaseUrl,
    [string]$ApiKey,
    [string]$SystemPrompt,
    [object]$InputObject,
    [string]$Model,
    [int]$TimeoutSec,
    [int]$MaxTokens,
    [switch]$UseSimpleUserPrompt,
    [switch]$DisableThinking,
    [hashtable]$RequestOptions
  )

  $inputJson = ConvertTo-LayerInputJson -InputObject $InputObject
  if ($UseSimpleUserPrompt) {
    $messages = @(
      @{
        role = "user"
        content = @"
你是视频质量裁判。只处理本次输入里的“当前一级标签”，只能从“当前二级标签定义”中选择标签。
请结合“原始提示词”和“本轮事实描述”描述错误点，再输出标签决策。
必须和“已有打标结果”去重，判断新增、替换、保留已有、移除已有、并列或不打标；尽量不并列。
证据不足时不打标。不要输出解释文字，只输出合法 JSON。

输出 JSON 结构：
{
  "当前一级标签": "",
  "本轮输入层级": "",
  "本批打标决策": [
    {
      "决策类型": "新增 | 替换 | 保留已有 | 移除已有 | 并列 | 不打标",
      "二级标签": "",
      "错因描述": "",
      "证据摘要": "",
      "匹配分数": 0,
      "置信度": "高 | 中 | 低",
      "相关已有标签编号": "",
      "处理说明": ""
    }
  ],
  "更新后打标结果": [
    {
      "标签编号": "",
      "一级标签": "",
      "二级标签": "",
      "错因描述": "",
      "证据来源": [
        {
          "层级": "",
          "位置": "",
          "证据摘要": ""
        }
      ],
      "匹配分数": 0,
      "置信度": "",
      "状态": "保留 | 新增 | 替换后保留 | 被替换 | 被移除"
    }
  ]
}

输入 JSON：
$inputJson
"@
      }
    )
  }
  else {
    $messages = @(
      @{
        role = "system"
        content = $SystemPrompt
      },
      @{
        role = "user"
        content = $inputJson
      }
    )
  }

  $body = ConvertTo-ChatBodyJson -Model $Model -Messages $messages -MaxTokens $MaxTokens -DisableThinking ([bool]$DisableThinking)

  $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
  return Invoke-JsonPostUtf8 `
    -Uri (Get-ChatCompletionsUrl -BaseUrl $BaseUrl) `
    -ApiKey $ApiKey `
    -BodyBytes $bytes `
    -TimeoutSec $TimeoutSec `
    -RequestOptions $RequestOptions
}

$apiConfig = Read-ApiConfig -Path $ApiKeyPath
if ([string]::IsNullOrWhiteSpace($apiConfig.Key) -or [string]::IsNullOrWhiteSpace($apiConfig.BaseUrl)) {
  throw "API config must contain key and url: $ApiKeyPath"
}

$requestOptions = @{}
if (-not [string]::IsNullOrWhiteSpace($Proxy)) {
  $requestOptions.Proxy = $Proxy
}

$videoAnalysisDir = Join-Path $AnalysisDir $VideoName
$targetDir = Join-Path $OutputDir $VideoName
New-Item -ItemType Directory -Force -Path $targetDir | Out-Null

$judgePrompt = Get-Content -Path (Join-Path $PromptDir $JudgePromptFile) -Raw -Encoding UTF8
$labels = Get-Content -Path (Join-Path $PromptDir "labels.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$originalPrompt = Get-Content -Path $OriginalPromptPath -Raw -Encoding UTF8

$macro = Get-Content -Path (Join-Path $videoAnalysisDir "macro.json") -Raw -Encoding UTF8
$segment = Get-Content -Path (Join-Path $videoAnalysisDir "segment.json") -Raw -Encoding UTF8
$frame = Get-Content -Path (Join-Path $videoAnalysisDir "frame.json") -Raw -Encoding UTF8

$existingLabels = @()
$batchIndex = 1
foreach ($category in $labels.PSObject.Properties) {
  $categoryName = $category.Name
  if ($Categories.Count -gt 0 -and $Categories -notcontains $categoryName) {
    continue
  }

  $layerInputs = @()
  if ($LayerMode -eq "Combined") {
    $layerInputs = @(
      @{
        Name = "三层合并"
        Object = [ordered]@{
          "原始提示词" = $originalPrompt
          "当前一级标签" = $categoryName
          "当前二级标签定义" = $category.Value
          "本轮输入层级" = "三层合并"
          "宏观层描述" = $macro
          "片段层描述" = $segment
          "逐帧层描述" = $frame
          "已有打标结果" = $existingLabels
        }
      }
    )
  }
  else {
    $layerInputs = @(
      @{
        Name = "宏观层"
        Object = [ordered]@{
          "原始提示词" = $originalPrompt
          "当前一级标签" = $categoryName
          "当前二级标签定义" = $category.Value
          "本轮输入层级" = "宏观层"
          "宏观层描述" = $macro
          "已有打标结果" = $existingLabels
        }
      },
      @{
        Name = "片段层"
        Object = [ordered]@{
          "原始提示词" = $originalPrompt
          "当前一级标签" = $categoryName
          "当前二级标签定义" = $category.Value
          "本轮输入层级" = "片段层"
          "片段层描述" = $segment
          "已有打标结果" = $existingLabels
        }
      },
      @{
        Name = "逐帧层"
        Object = [ordered]@{
          "原始提示词" = $originalPrompt
          "当前一级标签" = $categoryName
          "当前二级标签定义" = $category.Value
          "本轮输入层级" = "逐帧层"
          "逐帧层描述" = $frame
          "已有打标结果" = $existingLabels
        }
      }
    )
  }

  foreach ($layerInput in $layerInputs) {
    $layerName = $layerInput.Name
    if ($LayerNames.Count -gt 0 -and $LayerNames -notcontains $layerName) {
      continue
    }

    Write-Host "Judging $VideoName category: $categoryName layer: $layerName"

    try {
      $attempt = 0
      while ($true) {
        try {
          $attempt++
          $response = Invoke-Judge `
            -BaseUrl $apiConfig.BaseUrl `
            -ApiKey $apiConfig.Key `
            -SystemPrompt $judgePrompt `
            -InputObject $layerInput.Object `
            -Model $Model `
            -TimeoutSec $RequestTimeoutSec `
            -MaxTokens $MaxTokens `
            -UseSimpleUserPrompt:$UseSimpleUserPrompt `
            -DisableThinking:$DisableThinking `
            -RequestOptions $requestOptions
          break
        }
        catch {
          if ($attempt -gt $MaxRetries) {
            throw
          }
          Write-Host ("Judge request retry {0}/{1} for {2} / {3}: {4}" -f $attempt, $MaxRetries, $categoryName, $layerName, $_.Exception.Message)
          Start-Sleep -Seconds (3 * $attempt)
        }
      }
    }
    catch {
      Write-Host ("Judge request failed for {0} / {1}: {2}" -f $categoryName, $layerName, $_.Exception.Message)
      if ($ContinueOnBatchError) {
        $batchIndex++
        continue
      }
      throw
    }

    $safeLayerName = $layerName -replace '[\\/:*?"<>|]', '_'
    ConvertTo-Json -InputObject $response -Depth 80 | Set-Content -Path (Join-Path $targetDir ("{0:D2}_{1}_{2}_response.json" -f $batchIndex, $categoryName, $safeLayerName)) -Encoding UTF8
    $content = Get-ResponseText -Response $response
    $batchPath = Join-Path $targetDir ("{0:D2}_{1}_{2}.json" -f $batchIndex, $categoryName, $safeLayerName)
    $content | Set-Content -Path $batchPath -Encoding UTF8

    try {
      $parsed = ConvertFrom-ModelJson -Text $content
      if ($parsed."更新后打标结果") {
        $existingLabels = @($parsed."更新后打标结果")
        Save-FinalLabels -TargetDir $targetDir -Labels $existingLabels
      }
    }
    catch {
      Write-Host "Warning: judge output is not valid JSON for $categoryName / $layerName. Keeping previous labels."
    }

    $batchIndex++
  }
}

Save-FinalLabels -TargetDir $targetDir -Labels $existingLabels
Write-Host "Labels saved to: $targetDir"
