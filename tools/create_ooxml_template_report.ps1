$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$root = "C:\Users\10657\OneDrive\Desktop\Vedio2Text"
$template = "C:\Users\10657\OneDrive\Desktop\贸大模板\奶茶棕.pptx"
$outDir = Join-Path $root "汇报PPT"
$outPptx = Join-Path $outDir "分层证据驱动的视频生成质量评估框架_奶茶棕模板.pptx"

New-Item -ItemType Directory -Force -Path $outDir | Out-Null
Copy-Item -LiteralPath $template -Destination $outPptx -Force

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

function MaxShapeId([xml]$xml) {
  $nsm = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
  $nsm.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")
  $ids = $xml.SelectNodes("//p:cNvPr", $nsm) | ForEach-Object { [int]$_.id }
  if ($ids.Count -eq 0) { return 1000 }
  return (($ids | Measure-Object -Maximum).Maximum + 1)
}

function Apply-SlideOverlay {
  param([System.IO.Compression.ZipArchive]$zip, [int]$slideNo, [string[]]$shapes)
  $entryName = "ppt/slides/slide$slideNo.xml"
  $entry = $zip.GetEntry($entryName)
  $reader = New-Object IO.StreamReader($entry.Open())
  $xmlText = $reader.ReadToEnd()
  $reader.Close()
  $xmlText = [regex]::Replace(
    $xmlText,
    "<a:t>([^<]*(Add|ADD|Your|YOUR|Title|TITLE|99%|tecxt)[^<]*)</a:t>",
    "<a:t></a:t>"
  )
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
}

function Box($l,$t,$w,$h,$fill="FFF9F1",$alpha=0.96) {
  return AddShapeXml -id "__ID__" -name "report-panel" -left $l -top $t -width $w -height $h -fill $fill -alpha $alpha
}
function Txt($text,$l,$t,$w,$h,$size=28,$color="2B211B",$bold=$false,$align="l") {
  return AddShapeXml -id "__ID__" -name "report-text" -left $l -top $t -width $w -height $h -text $text -fontSize $size -fontColor $color -bold $bold -align $align
}
function Footer($page,$section) {
  return @(
    (Txt $section 92 1018 900 34 18 "765E4D" $false "l"),
    (Txt ("{0:D2} / 12" -f $page) 1700 1018 150 34 18 "765E4D" $false "r")
  )
}

$slides = @{}

$slides[1] = @(
  (Box 80 150 1320 620 "5B3D2B" 0.82),
  (Txt "分层证据驱动的视频生成质量评估框架" 130 235 1220 170 58 "FFF9F1" $true),
  (Box 135 445 260 6 "FFF9F1" 1),
  (Txt "多模态视频解析 · 三层证据链 · 标签裁判" 135 500 900 52 30 "FFF9F1"),
  (Txt "基于项目记录、Prompt 文档、样本视频、analysis 与 labels 结果生成" 135 570 1020 46 24 "FFF9F1")
) + (Footer 1 "METHOD REPORT")

$slides[2] = @(
  (Box 980 105 830 780 "FFF9F1" 0.96),
  (Txt "研究背景与问题定义" 1060 150 660 86 46 "5B3D2B" $true),
  (Box 1062 260 180 6 "845C3F" 1),
  (Txt "• 人工评价可信，但规模化成本高、口径维护困难`n• 端到端视频打标容易产生幻觉、遗漏与不可追溯结论`n• 视频生成质量评估需要跨主体、属性、动作、场景、时间、音频与安全多个维度精确归因" 1060 330 720 310 25 "261E19"),
  (Txt "核心转化：把复杂的多模态判断，拆解为可审计的文本证据判定。" 1060 730 700 100 30 "845C3F" $true)
) + (Footer 2 "BACKGROUND")

$slides[3] = @(
  (Box 90 90 1740 780 "FFF9F1" 0.94),
  (Txt "多模态视频解析的技术约束" 130 135 760 72 44 "5B3D2B" $true),
  (Box 130 285 500 235 "FFF9F1" 1),(Txt "抽帧限制" 164 315 420 46 30 "5B3D2B" $true),(Txt "视频模型通常处理采样帧或片段特征，不等同于逐帧穷举，因此短暂错误可能漏检。" 164 380 420 110 22 "765E4D"),
  (Box 690 285 500 235 "FFF9F1" 1),(Txt "表征压缩" 724 315 420 46 30 "5B3D2B" $true),(Txt "视觉编码器将连续画面压缩为视觉 token，细粒度文字、手部和局部结构容易弱化。" 724 380 420 110 22 "765E4D"),
  (Box 1250 285 500 235 "845C3F" 1),(Txt "语义幻觉" 1284 315 420 46 30 "FFF9F1" $true),(Txt "语言模型可能在证据不足时补全不存在的对象或场景，造成误报。" 1284 380 420 110 22 "FFF9F1"),
  (Txt "参考脉络：Flamingo、Video-ChatGPT、Video-LLaVA、POPE、MVBench、Video-MME" 132 610 980 45 20 "765E4D")
) + (Footer 3 "TECHNICAL PRINCIPLE")

$slides[4] = @(
  (Box 650 80 1180 760 "FFF9F1" 0.95),
  (Txt "系统流程：先描述，后裁判" 760 110 920 72 46 "5B3D2B" $true),
  (Txt "原始输入 → 三层描述 → 分批裁判 → 去重归因 → 最终 JSON" 760 245 980 56 32 "845C3F" $true),
  (Txt "01 原始输入：Vedio.mp4、Prompt.md、标签体系与各级提示词`n02 三层描述：宏观层、片段层、逐帧层只生成事实描述`n03 分批裁判：按一级标签输入二级标签定义与证据`n04 去重归因：判断新增、替换、保留或移除，减少重复打标`n05 最终 JSON：输出标签对、证据来源与错因描述" 760 350 940 310 25 "261E19")
) + (Footer 4 "PIPELINE")

$slides[5] = @(
  (Box 90 110 1740 760 "FFF9F1" 0.94),
  (Txt "三层证据权责" 135 145 760 80 50 "5B3D2B" $true),
  (Txt "描述层不可见标签，只负责把可观察事实记录清楚。" 138 245 880 50 26 "765E4D"),
  (Box 140 385 500 300 "FFF9F1" 1),(Txt "宏观层" 174 420 420 42 30 "5B3D2B" $true),(Txt "全片主体、主体属性、整体场景、视觉风格、全局事件、整体音频与敏感内容摘要。" 174 492 420 120 23 "765E4D"),
  (Box 710 385 500 300 "845C3F" 1),(Txt "片段层" 744 420 420 42 30 "FFF9F1" $true),(Txt "时间段、动作过程、事件顺序、相邻片段变化、接触关系与音画时间关系。" 744 492 420 120 23 "FFF9F1"),
  (Box 1280 385 500 300 "FFF9F1" 1),(Txt "逐帧层" 1314 420 420 42 30 "5B3D2B" $true),(Txt "单帧主体、局部结构、清晰度、画面文字、构图、色彩与相邻帧变化。" 1314 492 420 120 23 "765E4D")
) + (Footer 5 "THREE-LAYER EVIDENCE")

$slides[6] = @(
  (Box 730 80 1100 850 "FFF9F1" 0.95),
  (Txt "标签体系：7 个一级标签，31 个二级标签" 760 120 980 80 42 "5B3D2B" $true),
  (Txt "提示词一致性：主体、属性、动作、场景、风格`n视觉生成质量：清晰度、结构、局部失败、伪影`n时间一致性：角色、物体、场景、动作、事件`n物理真实性：人体、物体、碰撞、自然现象`n音频生成质量：台词、同步、失真、情绪、环境音`n审美质量：构图、镜头、色彩、节奏`n安全与合规：暴力、色情、隐私、版权" 850 260 800 520 26 "261E19")
) + (Footer 6 "LABEL ONTOLOGY")

$slides[7] = @(
  (Box 90 70 1740 870 "FFF9F1" 0.95),
  (Box 120 760 620 190 "FFF9F1" 1),
  (Box 650 760 1040 190 "FFF9F1" 1),
  (Txt "样本视频：真人古装写实短片" 170 95 920 78 46 "5B3D2B" $true),
  (Box 185 220 920 520 "261E19" 1),
  (Txt "VEDIO.MP4" 245 390 800 78 52 "FFF9F1" $true "ctr"),
  (Txt "视频素材路径：sample\Vedio.mp4" 245 480 800 45 26 "FFF9F1" $false "ctr"),
  (Txt "PPT 中保留视频入口说明；原始视频文件与本 PPT 位于同一项目目录。" 245 545 800 70 22 "FFF9F1" $false "ctr"),
  (Box 1165 230 540 390 "FFF9F1" 1),
  (Txt "核心提示词要素" 1205 270 420 48 32 "5B3D2B" $true),
  (Txt "• 清晨薄雾中的古代庭院`n• 年轻女书生，浅青交领长衫，木簪束发`n• 竹简出现清晰黑色文字《古风测试》`n• 水缸倒影、竹叶落下、动作连续、画面清晰" 1205 340 450 220 22 "261E19")
) + (Footer 7 "SAMPLE VIDEO")

$slides[8] = @(
  (Box 90 110 1740 760 "FFF9F1" 0.94),
  (Txt "证据抽取：提示词要素核对" 155 140 780 70 46 "5B3D2B" $true),
  (Txt "先从 Prompt 提取核心要素，再在三层描述中记录：有、无、不可见、未知。" 158 232 930 46 25 "765E4D"),
  (Box 150 360 500 270 "FFF9F1" 1),(Txt "宏观层" 184 392 420 40 30 "5B3D2B" $true),(Txt "确认古代庭院、服装、竹简等全局要素；指出未出现指定文字《古风测试》。" 184 460 420 110 22 "765E4D"),
  (Box 710 360 500 270 "FFF9F1" 1),(Txt "片段层" 744 392 420 40 30 "5B3D2B" $true),(Txt "按 0-3 秒、3-5 秒、5-8 秒描述动作、场景与道具变化。" 744 460 420 110 22 "765E4D"),
  (Box 1270 360 500 270 "845C3F" 1),(Txt "逐帧层" 1304 392 420 40 30 "FFF9F1" $true),(Txt "帧3：文字模糊不可读，头部被裁切；帧4-5：发型变化为辫子。" 1304 460 420 110 22 "FFF9F1")
) + (Footer 8 "EVIDENCE EXTRACTION")

$slides[9] = @(
  (Box 0 0 1920 1080 "5B3D2B" 1),
  (Txt "裁判输出：最终标签与错因" 118 105 920 76 46 "FFF9F1" $true),
  (Txt "最终结果只展示：一级标签--二级标签、证据来源、错因描述。" 120 205 980 46 25 "FFF9F1"),
  (Box 135 315 1580 580 "FFF9F1" 1),
  (Txt "提示词一致性--属性错误 | 逐帧层·帧3 | 竹简文字模糊不可读，未识别到指定字样《古风测试》。`n提示词一致性--动作错误 | 逐帧层·帧3-5 | 未出现左手按住毛笔动作，且桌面出现非指定剪刀类物件。`n时间一致性--角色突变 | 逐帧层·帧1-3 vs 帧4-5 | 人物发型从木簪束发无原因变化为辫子。`n审美质量--构图问题 | 逐帧层·帧3 | 人物头部顶部被画面上边缘裁切，影响主体完整性。" 170 360 1500 420 23 "261E19")
) + (Footer 9 "JUDGE RESULT")

$slides[10] = @(
  (Box 450 90 1400 780 "FFF9F1" 0.95),
  (Txt "关键改进：分批裁判与去重归因" 520 115 980 72 46 "5B3D2B" $true),
  (Box 520 285 420 250 "FFF9F1" 1),(Txt "逐一级标签分批" 554 322 350 40 28 "5B3D2B" $true),(Txt "每次只关注一个一级标签下的二级标签，降低上下文压力。" 554 390 340 90 22 "765E4D"),
  (Box 980 285 420 250 "845C3F" 1),(Txt "携带已有打标" 1014 322 350 40 28 "FFF9F1" $true),(Txt "下一批判断当前错因是否已存在，选择替换、保留或移除。" 1014 390 340 90 22 "FFF9F1"),
  (Box 1440 285 360 250 "FFF9F1" 1),(Txt "紧凑最终输出" 1474 322 290 40 28 "5B3D2B" $true),(Txt "中间过程可分层、分批；最终汇总为一个 JSON 文件。" 1474 390 280 90 22 "765E4D"),
  (Txt "目标：减少模型反复归因，提高标签结果接近人工标注中的主因判断。" 520 625 980 60 30 "845C3F" $true)
) + (Footer 10 "IMPROVEMENT")

$slides[11] = @(
  (Box 840 90 900 820 "FFF9F1" 0.95),
  (Txt "局限与后续迭代：自适应采样" 890 135 820 76 42 "5B3D2B" $true),
  (Txt "• 逐帧层不是对 24 FPS 视频的全帧穷举`n• 均匀抽帧可能漏掉手部瞬时畸变、物体短暂穿模、文字一帧错误`n• 更可行的策略是：均匀采样 + 变化点采样 + 异常片段加密采样`n• 当裁判证据不足或低置信时，回到对应时间段补充采样" 890 280 820 430 25 "261E19"),
  (Box 930 760 720 90 "845C3F" 1),
  (Txt "优先选择画面骤变帧、不连贯帧和已提示异常的时间段" 970 788 640 42 25 "FFF9F1" $true)
) + (Footer 11 "LIMITATION")

$slides[12] = @(
  (Box 80 150 1320 700 "5B3D2B" 0.82),
  (Txt "从视频错误到可追溯证据链" 130 250 1200 120 58 "FFF9F1" $true),
  (Box 134 415 260 6 "FFF9F1" 1),
  (Txt "• 描述层负责事实，避免提前带入标签结论`n• 裁判层负责归因，结合 Prompt、标签定义和三层证据`n• 数据看板负责迭代，持续分析准确率、召回率与重复归因" 135 485 1050 250 30 "FFF9F1")
) + (Footer 12 "END")

$zip = [System.IO.Compression.ZipFile]::Open($outPptx, [System.IO.Compression.ZipArchiveMode]::Update)
try {
  foreach ($k in ($slides.Keys | Sort-Object)) {
    Apply-SlideOverlay $zip $k $slides[$k]
  }
} finally {
  $zip.Dispose()
}

Write-Output $outPptx







