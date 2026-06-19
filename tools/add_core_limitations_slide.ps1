$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$root = "C:\Users\10657\OneDrive\Desktop\Vedio2Text"
$pptx = Join-Path $root "汇报PPT\分层证据驱动的视频生成质量评估框架_奶茶棕模板.pptx"

$slideW = 24379238.0
$slideH = 13716000.0

function To-X([double]$px) { return [int][Math]::Round($px / 1920.0 * $slideW) }
function To-Y([double]$px) { return [int][Math]::Round($px / 1080.0 * $slideH) }
function To-W([double]$px) { return [int][Math]::Round($px / 1920.0 * $slideW) }
function To-H([double]$px) { return [int][Math]::Round($px / 1080.0 * $slideH) }

function Escape-Xml([string]$s) {
  return [System.Security.SecurityElement]::Escape($s)
}

function TextParagraphXml([string]$text, [int]$fontSize, [string]$color, [bool]$bold = $false, [string]$align = "l") {
  $b = if ($bold) { ' b="1"' } else { "" }
  $sz = $fontSize * 100
  $runs = @()
  foreach ($line in ($text -split "`n")) {
    $escaped = Escape-Xml $line
    $runs += "<a:p><a:pPr algn=`"$align`"/><a:r><a:rPr lang=`"zh-CN`" sz=`"$sz`"$b dirty=`"0`"><a:solidFill><a:srgbClr val=`"$color`"/></a:solidFill><a:latin typeface=`"Microsoft YaHei UI`"/><a:ea typeface=`"Microsoft YaHei UI`"/><a:cs typeface=`"Microsoft YaHei UI`"/></a:rPr><a:t>$escaped</a:t></a:r></a:p>"
  }
  return ($runs -join "")
}

function AddShapeXml {
  param(
    [string]$id,
    [string]$name,
    [double]$left,
    [double]$top,
    [double]$width,
    [double]$height,
    [string]$fill = "",
    [double]$alpha = 1.0,
    [string]$text = "",
    [int]$fontSize = 24,
    [string]$fontColor = "2B211B",
    [bool]$bold = $false,
    [string]$align = "l"
  )

  $x = To-X $left; $y = To-Y $top; $cx = To-W $width; $cy = To-H $height
  $fillXml = "<a:noFill/>"
  if ($fill -ne "") {
    $alphaVal = [int][Math]::Round($alpha * 100000)
    $alphaXml = if ($alphaVal -lt 100000) { "<a:alpha val=`"$alphaVal`"/>" } else { "" }
    $fillXml = "<a:solidFill><a:srgbClr val=`"$fill`">$alphaXml</a:srgbClr></a:solidFill>"
  }
  $txBox = if ($text -ne "") { ' txBox="1"' } else { "" }
  $txBody = ""
  if ($text -ne "") {
    $paras = TextParagraphXml $text $fontSize $fontColor $bold $align
    $txBody = "<p:txBody><a:bodyPr wrap=`"square`" rtlCol=`"0`" anchor=`"t`"><a:spAutoFit/></a:bodyPr><a:lstStyle/>$paras</p:txBody>"
  }

  return @"
<p:sp>
  <p:nvSpPr><p:cNvPr id="$id" name="$name"/><p:cNvSpPr$txBox/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="$x" y="$y"/><a:ext cx="$cx" cy="$cy"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>$fillXml<a:ln><a:noFill/></a:ln></p:spPr>
  $txBody
</p:sp>
"@
}

function Box($l,$t,$w,$h,$fill="FFF9F1",$alpha=0.96) {
  return AddShapeXml -id "__ID__" -name "core-limitations-panel" -left $l -top $t -width $w -height $h -fill $fill -alpha $alpha
}

function Txt($text,$l,$t,$w,$h,$size=28,$color="2B211B",$bold=$false,$align="l") {
  return AddShapeXml -id "__ID__" -name "core-limitations-text" -left $l -top $t -width $w -height $h -text $text -fontSize $size -fontColor $color -bold $bold -align $align
}

function MaxShapeId([xml]$xml) {
  $nsm = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
  $nsm.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")
  $ids = $xml.SelectNodes("//p:cNvPr", $nsm) | ForEach-Object { [int]$_.id }
  if ($ids.Count -eq 0) { return 1000 }
  return (($ids | Measure-Object -Maximum).Maximum + 1)
}

$slideNo = 12
$entryName = "ppt/slides/slide$slideNo.xml"

$shapes = @(
  (Box 0 0 1920 1080 "5B3D2B" 0.86),
  (Box 96 86 1728 900 "FFF9F1" 0.98),
  (Txt "当前项目的核心局限" 145 125 920 72 46 "5B3D2B" $true),
  (Box 145 218 240 6 "845C3F" 1),
  (Txt "当前链路已经形成可运行闭环，但仍有三类问题会限制标签结果的稳定性与扩展上限。" 145 250 1320 46 24 "765E4D"),
  (Box 145 360 500 360 "FFF9F1" 1),
  (Txt "01  主观维度仍弱" 185 400 420 44 30 "5B3D2B" $true),
  (Txt "审美、灵动感、氛围、观看体验等标签高度依赖人工偏好和任务语境，仅靠提示词工程难以稳定提升一致性。后续需要更细粒度评价量表或人工标注数据支撑。" 185 470 410 145 22 "765E4D"),
  (Box 710 360 500 360 "845C3F" 1),
  (Txt "02  细粒度画面仍可能漏检" 750 400 420 44 30 "FFF9F1" $true),
  (Txt "逐帧层本质是关键帧或高频采样帧描述，不是对原始视频的全帧穷举。手部瞬时畸变、文字一帧错误、短暂穿模等问题仍可能被抽帧策略漏掉。" 750 470 410 145 22 "FFF9F1"),
  (Box 1275 360 500 360 "FFF9F1" 1),
  (Txt "03  证据质量会向后传导" 1315 400 420 44 30 "5B3D2B" $true),
  (Txt "裁判模型依赖三层文字证据。如果分析层遗漏、幻觉或表述不清，后续打标会继承这些误差。因此仍需要提示词要素核对、低置信复核和人工抽检机制。" 1315 470 410 145 22 "765E4D"),
  (Box 145 785 1410 92 "5B3D2B" 1),
  (Txt "后续方向：主观量表细化 + 自适应采样 + 低置信证据回填 + 人工复核闭环" 185 812 1320 38 28 "FFF9F1" $true),
  (Txt "CORE LIMITATIONS" 92 1018 900 34 18 "765E4D"),
  (Txt "12 / 15" 1700 1018 150 34 18 "765E4D" $false "r")
)

$zip = [System.IO.Compression.ZipFile]::Open($pptx, [System.IO.Compression.ZipArchiveMode]::Update)
try {
  $entry = $zip.GetEntry($entryName)
  $reader = New-Object IO.StreamReader($entry.Open())
  $xmlText = $reader.ReadToEnd()
  $reader.Close()
  $xmlText = [regex]::Replace($xmlText, "<a:t>([^<]*(Add|ADD|Your|YOUR|Title|TITLE|99%|tecxt)[^<]*)</a:t>", "<a:t></a:t>")
  [xml]$xml = $xmlText
  $nextId = MaxShapeId $xml
  $shapeXml = ""
  foreach ($shape in $shapes) {
    $shapeXml += ($shape -replace "__ID__", [string]$nextId)
    $nextId++
  }
  $newXml = $xmlText -replace "</p:spTree>", ($shapeXml + "</p:spTree>")
  $entry.Delete()
  $newEntry = $zip.CreateEntry($entryName)
  $writer = New-Object IO.StreamWriter($newEntry.Open(), [System.Text.UTF8Encoding]::new($false))
  $writer.Write($newXml)
  $writer.Close()
} finally {
  $zip.Dispose()
}

Write-Output $pptx

