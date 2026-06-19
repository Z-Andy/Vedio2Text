param(
  [string]$ApiKeyPath = "..\apikey1.txt",
  [string]$PromptPath = ".\video_generation_prompt.md",
  [string]$JobPath = ".\video_job.json",
  [string]$StatusPath = ".\video_status.json",
  [string]$OutputPath = ".\generated_video.mp4",
  [string]$Model = "sora-2",
  [string]$Size = "1280x720",
  [int]$Seconds = 8,
  [int]$PollIntervalSeconds = 10,
  [int]$MaxPollCount = 120,
  [string]$Proxy = "",
  [switch]$NoSize,
  [int]$SubmitRetryCount = 5,
  [int]$SubmitRetryDelaySeconds = 30
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Read-VideoPrompt {
  param([string]$Path)

  $content = Get-Content -Path $Path -Raw -Encoding UTF8
  $promptHeading = "## " + [string][char]0x63D0 + [string][char]0x793A + [string][char]0x8BCD
  $parts = $content -split [regex]::Escape($promptHeading), 2
  if ($parts.Count -ge 2) {
    return $parts[1].Trim()
  }

  return $content.Trim()
}

function Read-ApiConfig {
  param([string]$Path)

  $raw = (Get-Content -Path $Path -Raw -Encoding UTF8).Trim()
  if ([string]::IsNullOrWhiteSpace($raw)) {
    throw "API config is empty: $Path"
  }

  try {
    $json = $raw | ConvertFrom-Json
    return @{
      Key = [string]$json.key
      BaseUrl = ([string]$json.url).TrimEnd("/")
    }
  }
  catch {
    try {
      $json = ("{" + $raw.Trim().Trim(",") + "}") | ConvertFrom-Json
      return @{
        Key = [string]$json.key
        BaseUrl = ([string]$json.url).TrimEnd("/")
      }
    }
    catch {
    }

    return @{
      Key = $raw
      BaseUrl = "https://api.openai.com"
    }
  }
}

function New-MultipartFormData {
  param([hashtable]$Fields)

  $boundary = "----Vedio2TextBoundary" + [guid]::NewGuid().ToString("N")
  $encoding = [System.Text.Encoding]::UTF8
  $stream = New-Object System.IO.MemoryStream

  foreach ($name in $Fields.Keys) {
    $value = [string]$Fields[$name]
    $part = "--$boundary`r`n" +
      "Content-Disposition: form-data; name=`"$name`"`r`n`r`n" +
      "$value`r`n"
    $bytes = $encoding.GetBytes($part)
    $stream.Write($bytes, 0, $bytes.Length)
  }

  $end = "--$boundary--`r`n"
  $endBytes = $encoding.GetBytes($end)
  $stream.Write($endBytes, 0, $endBytes.Length)

  return @{
    Body = $stream.ToArray()
    ContentType = "multipart/form-data; boundary=$boundary"
  }
}

function Get-ErrorText {
  param($ErrorRecord)

  $texts = @()
  if ($ErrorRecord.Exception.Message) {
    $texts += [string]$ErrorRecord.Exception.Message
  }

  try {
    $response = $ErrorRecord.Exception.Response
    if ($null -ne $response) {
      $stream = $response.GetResponseStream()
      if ($null -ne $stream) {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $texts += $reader.ReadToEnd()
        $reader.Close()
      }
    }
  }
  catch {
  }

  return ($texts -join "`n")
}

$apiConfig = Read-ApiConfig -Path $ApiKeyPath
$apiKey = $apiConfig.Key
$baseUrl = $apiConfig.BaseUrl

if ([string]::IsNullOrWhiteSpace($apiKey)) {
  throw "API key is empty: $ApiKeyPath"
}

if ([string]::IsNullOrWhiteSpace($baseUrl)) {
  throw "API base URL is empty: $ApiKeyPath"
}

$prompt = Read-VideoPrompt -Path $PromptPath
$headers = @{
  Authorization = "Bearer $apiKey"
}

$requestOptions = @{}
if (-not [string]::IsNullOrWhiteSpace($Proxy)) {
  $requestOptions.Proxy = $Proxy
}

$createFields = @{
  model = $Model
  prompt = $prompt
  seconds = [string]$Seconds
}
if (-not $NoSize -and -not [string]::IsNullOrWhiteSpace($Size)) {
  $createFields.size = $Size
}

$multipart = New-MultipartFormData -Fields $createFields

$job = $null
for ($submitAttempt = 1; $submitAttempt -le $SubmitRetryCount; $submitAttempt++) {
  try {
    Write-Host "Creating video generation job via $baseUrl ... attempt $submitAttempt/$SubmitRetryCount"
    $job = Invoke-RestMethod `
      -Uri "$baseUrl/v1/videos" `
      -Method Post `
      -Headers $headers `
      -ContentType $multipart.ContentType `
      -Body $multipart.Body `
      @requestOptions
    break
  }
  catch {
    $message = Get-ErrorText -ErrorRecord $_
    if ($submitAttempt -ge $SubmitRetryCount -or ($message -notmatch "负载已饱和|load|busy|rate|saturated")) {
      throw
    }

    Write-Host "Submit failed because upstream is busy. Waiting $SubmitRetryDelaySeconds seconds before retry..."
    Start-Sleep -Seconds $SubmitRetryDelaySeconds
  }
}

$job | ConvertTo-Json -Depth 20 | Set-Content -Path $JobPath -Encoding UTF8
$videoId = $job.id
if ([string]::IsNullOrWhiteSpace($videoId)) {
  throw "No id found in create response. Response saved to: $JobPath"
}

Write-Host "Job created: $videoId"

$status = $null
for ($i = 1; $i -le $MaxPollCount; $i++) {
  Start-Sleep -Seconds $PollIntervalSeconds

  $status = Invoke-RestMethod `
    -Uri "$baseUrl/v1/videos/$videoId" `
    -Method Get `
    -Headers @{ Authorization = "Bearer $apiKey" } `
    @requestOptions

  $status | ConvertTo-Json -Depth 20 | Set-Content -Path $StatusPath -Encoding UTF8
  Write-Host "Poll $i`: $($status.status)"

  if ($status.status -eq "completed" -or $status.status -eq "succeeded") {
    break
  }

  if ($status.status -eq "failed" -or $status.status -eq "cancelled") {
    throw "Video job did not complete. Status: $($status.status). Details saved to: $StatusPath"
  }
}

if ($null -eq $status -or ($status.status -ne "completed" -and $status.status -ne "succeeded")) {
  throw "Polling timed out. Last status saved to: $StatusPath"
}

Write-Host "Downloading video..."
Invoke-WebRequest `
  -Uri "$baseUrl/v1/videos/$videoId/content" `
  -Method Get `
  -Headers @{ Authorization = "Bearer $apiKey" } `
  -OutFile $OutputPath `
  -UseBasicParsing `
  @requestOptions

Write-Host "Video saved: $OutputPath"
