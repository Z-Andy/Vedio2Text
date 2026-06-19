param(
  [string]$ApiKeyPath = "..\apikey2.txt",
  [string]$VideoPath = ".\Vedio1.mp4",
  [string]$OriginalPromptPath = "",
  [string]$OutputDir = ".\analysis",
  [string]$PromptDir = "..\Prompt",
  [string]$Model = "doubao-seed-2-0-lite-260428",
  [string]$Proxy = ""
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Read-ApiConfig {
  param([string]$Path)

  $raw = (Get-Content -Path $Path -Raw -Encoding UTF8).Trim()
  if ([string]::IsNullOrWhiteSpace($raw)) {
    throw "API config is empty: $Path"
  }

  if ($raw -match "Bearer\s+([A-Za-z0-9._\-]+)") {
    return @{
      Key = [string]$Matches[1]
      BaseUrl = "https://ark.cn-beijing.volces.com/api/v3"
      ApiStyle = "responses"
    }
  }

  try {
    $json = $raw | ConvertFrom-Json
  }
  catch {
    try {
      $json = ("{" + $raw.Trim().Trim(",") + "}") | ConvertFrom-Json
    }
    catch {
      if ($raw -like "ark-*") {
        return @{
          Key = $raw
          BaseUrl = "https://ark.cn-beijing.volces.com/api/v3"
          ApiStyle = "responses"
        }
      }

      return @{
        Key = $raw
        BaseUrl = "https://api.openai.com/v1"
        ApiStyle = "chat"
      }
    }
  }

  return @{
    Key = [string]$json.key
    BaseUrl = ([string]$json.url).TrimEnd("/")
    ApiStyle = "chat"
  }
}

function Get-ResponseText {
  param($Response)

  if ($Response.choices -and $Response.choices.Count -gt 0) {
    return [string]$Response.choices[0].message.content
  }

  if ($Response.output_text) {
    return [string]$Response.output_text
  }

  if ($Response.output -and $Response.output.Count -gt 0) {
    $texts = @()
    foreach ($item in $Response.output) {
      if ($item.content) {
        foreach ($content in $item.content) {
          if ($content.text) {
            $texts += [string]$content.text
          }
        }
      }
    }
    if ($texts.Count -gt 0) {
      return ($texts -join "`n")
    }
  }

  return ($Response | ConvertTo-Json -Depth 30)
}

function Remove-MarkdownCodeFence {
  param([string]$Text)

  $clean = $Text.Trim()
  if ($clean -match '^\s*```') {
    $clean = $clean -replace '^\s*```[A-Za-z0-9_-]*\s*', ''
    $clean = $clean -replace '\s*```\s*$', ''
  }

  return $clean.Trim()
}

function Get-ChatCompletionsUrl {
  param([string]$BaseUrl)

  $base = $BaseUrl.TrimEnd("/")
  if ($base -match "/(v1|api/v3)$") {
    return "$base/chat/completions"
  }

  return "$base/v1/chat/completions"
}

function Get-ResponsesUrl {
  param([string]$BaseUrl)

  $base = $BaseUrl.TrimEnd("/")
  if ($base -match "/(v1|api/v3)$") {
    return "$base/responses"
  }

  return "$base/v1/responses"
}

function Invoke-VideoAnalysis {
  param(
    [string]$BaseUrl,
    [string]$ApiKey,
    [string]$PromptText,
    [string]$VideoDataUrl,
    [string]$Model,
    [string]$ApiStyle,
    [hashtable]$RequestOptions
  )

  if ($ApiStyle -eq "responses") {
    $body = @{
      model = $Model
      input = @(
        @{
          role = "user"
          content = @(
            @{
              type = "input_video"
              video_url = $VideoDataUrl
            },
            @{
              type = "input_text"
              text = $PromptText
            }
          )
        }
      )
      temperature = 0
    } | ConvertTo-Json -Depth 50 -Compress

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
    return Invoke-RestMethod `
      -Uri (Get-ResponsesUrl -BaseUrl $BaseUrl) `
      -Method Post `
      -Headers @{ Authorization = "Bearer $ApiKey"; "Content-Type" = "application/json" } `
      -Body $bytes `
      @RequestOptions
  }

  $body = @{
    model = $Model
    messages = @(
      @{
        role = "user"
        content = @(
          @{
            type = "text"
            text = $PromptText
          },
          @{
            type = "video_url"
            video_url = @{
              url = $VideoDataUrl
            }
          }
        )
      }
    )
    temperature = 0
  } | ConvertTo-Json -Depth 50 -Compress

  $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
  return Invoke-RestMethod `
    -Uri (Get-ChatCompletionsUrl -BaseUrl $BaseUrl) `
    -Method Post `
    -Headers @{ Authorization = "Bearer $ApiKey"; "Content-Type" = "application/json" } `
    -Body $bytes `
    @RequestOptions
}

$apiConfig = Read-ApiConfig -Path $ApiKeyPath
if ([string]::IsNullOrWhiteSpace($apiConfig.Key) -or [string]::IsNullOrWhiteSpace($apiConfig.BaseUrl)) {
  throw "API config must contain key and url: $ApiKeyPath"
}

$videoFile = Get-Item -LiteralPath $VideoPath
$videoName = [System.IO.Path]::GetFileNameWithoutExtension($videoFile.Name)
$targetDir = Join-Path $OutputDir $videoName
New-Item -ItemType Directory -Force -Path $targetDir | Out-Null

if ([string]::IsNullOrWhiteSpace($OriginalPromptPath)) {
  $index = [regex]::Match($videoName, "\d+").Value
  if (-not [string]::IsNullOrWhiteSpace($index)) {
    $candidatePromptPath = ".\Prompt$index.md"
    if (Test-Path -LiteralPath $candidatePromptPath) {
      $OriginalPromptPath = $candidatePromptPath
    }
  }
}

$originalPrompt = ""
if (-not [string]::IsNullOrWhiteSpace($OriginalPromptPath) -and (Test-Path -LiteralPath $OriginalPromptPath)) {
  $originalPrompt = Get-Content -Path $OriginalPromptPath -Raw -Encoding UTF8
}

$requestOptions = @{}
if (-not [string]::IsNullOrWhiteSpace($Proxy)) {
  $requestOptions.Proxy = $Proxy
}

$videoBytes = [System.IO.File]::ReadAllBytes($videoFile.FullName)
$videoBase64 = [Convert]::ToBase64String($videoBytes)
$videoDataUrl = "data:video/mp4;base64,$videoBase64"

$layers = @(
  @{ Name = "macro"; File = "macro_prompt.md" },
  @{ Name = "segment"; File = "segment_prompt.md" },
  @{ Name = "frame"; File = "frame_prompt.md" }
)

foreach ($layer in $layers) {
  $promptPath = Join-Path $PromptDir $layer.File
  $layerPromptText = Get-Content -Path $promptPath -Raw -Encoding UTF8
  $promptText = @"
$layerPromptText

---

## 原始视频生成提示词

$originalPrompt
"@
  Write-Host "Analyzing $videoName with $($layer.Name) prompt..."

  $response = Invoke-VideoAnalysis `
    -BaseUrl $apiConfig.BaseUrl `
    -ApiKey $apiConfig.Key `
    -PromptText $promptText `
    -VideoDataUrl $videoDataUrl `
    -Model $Model `
    -ApiStyle $apiConfig.ApiStyle `
    -RequestOptions $requestOptions

  $response | ConvertTo-Json -Depth 50 | Set-Content -Path (Join-Path $targetDir "$($layer.Name)_response.json") -Encoding UTF8
  Remove-MarkdownCodeFence -Text (Get-ResponseText -Response $response) | Set-Content -Path (Join-Path $targetDir "$($layer.Name).json") -Encoding UTF8
}

Write-Host "Analysis saved to: $targetDir"
