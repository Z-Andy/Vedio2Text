$ErrorActionPreference = "Stop"

$root = "C:\Users\10657\OneDrive\Desktop\Vedio2Text"
$template = "C:\Users\10657\OneDrive\Desktop\贸大模板\奶茶棕.pptx"
$outDir = Join-Path $root "汇报PPT"
$outPptx = Join-Path $outDir "分层证据驱动的视频生成质量评估框架_奶茶棕模板.pptx"
$videoPath = Join-Path $root "sample\Vedio.mp4"

New-Item -ItemType Directory -Force -Path $outDir | Out-Null
Copy-Item -LiteralPath $template -Destination $outPptx -Force

$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $pp.Presentations.Open($outPptx, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

$slideW = $pres.PageSetup.SlideWidth
$slideH = $pres.PageSetup.SlideHeight

$msoTextBox = 17
$msoPlaceholder = 14
$msoFalse = [Microsoft.Office.Core.MsoTriState]::msoFalse
$msoTrue = [Microsoft.Office.Core.MsoTriState]::msoTrue
$ppSaveAsOpenXmlPresentation = 24

function Rgb {
  param([int]$r, [int]$g, [int]$b)
  return ($r -bor ($g -shl 8) -bor ($b -shl 16))
}

$brown = Rgb 91 61 43
$brown2 = Rgb 132 92 63
$beige = Rgb 242 229 211
$cream = Rgb 255 249 241
$white = Rgb 255 255 255
$black = Rgb 38 30 25
$muted = Rgb 118 94 77
$blue = Rgb 37 78 125

function Try-Com {
  param([scriptblock]$block)
  for ($i = 0; $i -lt 6; $i++) {
    try {
      & $block
      return
    } catch {
      Start-Sleep -Milliseconds 350
    }
  }
}

function Clear-SlideText {
  param($slide)
  if ($null -eq $slide -or $null -eq $slide.Shapes) { return }
  for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
    try {
      $shape = $slide.Shapes.Item($i)
      if ($null -eq $shape) { continue }
      if ($shape.Type -eq $msoTextBox) {
        $shape.Delete()
      } elseif ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
        $txt = $shape.TextFrame.TextRange.Text
        if ($txt -match "Add|Your|Title|text|99%|TITLE|tecxt") {
          $shape.TextFrame.TextRange.Text = ""
        }
      }
    } catch {}
  }
}

function Add-Box {
  param($slide, [double]$left, [double]$top, [double]$width, [double]$height, [int]$fillColor, [double]$transparency = 0)
  $shape = $null
  for ($i = 0; $i -lt 8 -and $null -eq $shape; $i++) {
    try { $shape = $slide.Shapes.AddShape(1, $left, $top, $width, $height) } catch { Start-Sleep -Milliseconds 350 }
  }
  if ($null -eq $shape) { return $null }
  try { $shape.Fill.ForeColor.RGB = $fillColor } catch { try { $shape.Fill.ForeColor.SchemeColor = 1 } catch {} }
  $shape.Fill.Transparency = $transparency
  $shape.Line.Visible = $msoFalse
  return $shape
}

function Add-Line {
  param($slide, [double]$left, [double]$top, [double]$width, [int]$color = $brown2)
  $shape = $null
  for ($i = 0; $i -lt 8 -and $null -eq $shape; $i++) {
    try { $shape = $slide.Shapes.AddShape(1, $left, $top, $width, 5) } catch { Start-Sleep -Milliseconds 350 }
  }
  if ($null -eq $shape) { return $null }
  try { $shape.Fill.ForeColor.RGB = $color } catch {}
  $shape.Line.Visible = $msoFalse
  return $shape
}

function Add-Text {
  param(
    $slide,
    [string]$text,
    [double]$left,
    [double]$top,
    [double]$width,
    [double]$height,
    [double]$fontSize,
    [int]$color = $black,
    [bool]$bold = $false,
    [string]$font = "Microsoft YaHei UI",
    [int]$align = 1,
    [double]$lineSpace = 1.08
  )
  if ($null -eq $slide) { return $null }
  $shape = $null
  for ($i = 0; $i -lt 10 -and $null -eq $shape; $i++) {
    try { $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height) } catch { Start-Sleep -Milliseconds 400 }
  }
  if ($null -eq $shape) { return $null }
  Try-Com { $shape.TextFrame.MarginLeft = 0 }
  Try-Com { $shape.TextFrame.MarginRight = 0 }
  Try-Com { $shape.TextFrame.MarginTop = 0 }
  Try-Com { $shape.TextFrame.MarginBottom = 0 }
  Try-Com { $shape.TextFrame.WordWrap = $msoTrue }
  Try-Com { $shape.TextFrame.AutoSize = 0 }
  Try-Com { $shape.TextFrame.TextRange.Text = $text }
  $range = $shape.TextFrame.TextRange
  try { $range.Font.Name = $font } catch {}
  try { $range.Font.Size = $fontSize } catch {}
  try { $range.Font.Color.RGB = $color } catch {}
  try { if ($bold) { $range.Font.Bold = $msoTrue } else { $range.Font.Bold = $msoFalse } } catch {}
  try { $range.ParagraphFormat.Alignment = $align } catch {}
  try { $range.ParagraphFormat.SpaceWithin = $lineSpace } catch {}
  return $shape
}

function Add-BulletList {
  param($slide, [string[]]$items, [double]$left, [double]$top, [double]$width, [double]$height, [double]$fontSize = 28, [int]$color = $black)
  $text = ($items | ForEach-Object { "• $_" }) -join "`r"
  $shape = Add-Text $slide $text $left $top $width $height $fontSize $color $false "Microsoft YaHei UI" 1 1.15
  return $shape
}

function Add-Footer {
  param($slide, [int]$page, [string]$section = "分层证据驱动的视频生成质量评估")
  Add-Text $slide $section 92 1018 900 34 18 $muted $false "Microsoft YaHei UI" | Out-Null
  Add-Text $slide ("{0:D2} / 12" -f $page) 1700 1018 150 34 18 $muted $false "Consolas" 3 | Out-Null
}

function Add-Card {
  param($slide, [string]$title, [string]$body, [double]$left, [double]$top, [double]$width, [double]$height, [bool]$accent = $false)
  $fill = $cream
  $titleColor = $brown
  $bodyColor = $muted
  if ($accent) {
    $fill = $brown2
    $titleColor = $white
    $bodyColor = $cream
  }
  Add-Box $slide $left $top $width $height $fill | Out-Null
  Add-Text $slide $title ($left + 34) ($top + 30) ($width - 68) 54 30 $titleColor $true | Out-Null
  Add-Text $slide $body ($left + 34) ($top + 98) ($width - 68) ($height - 128) 23 $bodyColor $false | Out-Null
}

# 保留模板原有版式与装饰，不破坏性删除页面元素；正式内容以半透明内容区覆盖占位文案。

# 1 Cover
$s = $pres.Slides.Item(1)
Add-Text $s "分层证据驱动的视频生成质量评估框架" 130 235 1220 250 70 $cream $true "Microsoft YaHei UI" | Out-Null
Add-Line $s 134 525 260 $cream | Out-Null
Add-Text $s "多模态视频解析 · 三层证据链 · 标签裁判" 135 575 900 52 34 $cream $false | Out-Null
Add-Text $s "基于项目记录、Prompt 文档、样本视频、analysis 与 labels 结果生成" 135 645 1020 46 26 $cream $false | Out-Null
Add-Footer $s 1 "METHOD REPORT"

# 2 Background
$s = $pres.Slides.Item(2)
Add-Text $s "研究背景与问题定义" 1060 150 660 86 54 $brown $true | Out-Null
Add-Line $s 1062 260 180 $brown2 | Out-Null
Add-BulletList $s @(
  "人工评价可信，但规模化成本高、口径维护困难",
  "端到端视频打标容易产生幻觉、遗漏与不可追溯结论",
  "视频生成质量评估需要跨主体、属性、动作、场景、时间、音频与安全多个维度精确归因"
) 1060 330 720 310 28 $black | Out-Null
Add-Text $s "核心转化：把复杂的多模态判断，拆解为可审计的文本证据判定。" 1060 730 700 100 34 $brown2 $true | Out-Null
Add-Footer $s 2 "BACKGROUND"

# 3 Technical constraints
$s = $pres.Slides.Item(3)
Add-Text $s "多模态视频解析的技术约束" 130 135 760 72 50 $brown $true | Out-Null
Add-Card $s "抽帧限制" "视频模型通常处理采样帧或片段特征，不等同于逐帧穷举，因此短暂错误可能漏检。" 130 285 500 235
Add-Card $s "表征压缩" "视觉编码器将连续画面压缩为视觉 token，细粒度文字、手部和局部结构容易弱化。" 690 285 500 235
Add-Card $s "语义幻觉" "语言模型可能在证据不足时补全不存在的对象或场景，造成误报。" 1250 285 500 235 $true
Add-Text $s "参考脉络：Flamingo、Video-ChatGPT、Video-LLaVA、POPE、MVBench、Video-MME" 132 610 980 45 22 $muted | Out-Null
Add-Footer $s 3 "TECHNICAL PRINCIPLE"

# 4 Pipeline
$s = $pres.Slides.Item(4)
Add-Text $s "系统流程：先描述，后裁判" 760 110 920 72 52 $brown $true | Out-Null
$xs = @(145, 485, 825, 1165, 1505)
$titles = @("原始输入", "三层描述", "分批裁判", "去重归因", "最终 JSON")
$bodies = @(
  "Vedio.mp4、Prompt.md、标签体系与各级提示词",
  "宏观层、片段层、逐帧层只生成事实描述",
  "按一级标签分批输入二级标签定义与证据",
  "判断新增、替换、保留或移除，减少重复打标",
  "输出标签对、证据来源与错因描述"
)
for ($i = 0; $i -lt 5; $i++) {
  Add-Box $s $xs[$i] 360 260 210 $(if($i -eq 1 -or $i -eq 4){$brown2}else{$cream}) | Out-Null
  $tc = $(if($i -eq 1 -or $i -eq 4){$white}else{$brown})
  $bc = $(if($i -eq 1 -or $i -eq 4){$cream}else{$muted})
  Add-Text $s ("0{0}" -f ($i+1)) ($xs[$i]+24) 384 80 35 22 $tc $false "Consolas" | Out-Null
  Add-Text $s $titles[$i] ($xs[$i]+24) 430 210 42 28 $tc $true | Out-Null
  Add-Text $s $bodies[$i] ($xs[$i]+24) 492 210 80 20 $bc $false | Out-Null
  if ($i -lt 4) {
    Add-Text $s "→" ($xs[$i]+275) 435 50 50 38 $brown2 $true | Out-Null
  }
}
Add-Footer $s 4 "PIPELINE"

# 5 Three levels
$s = $pres.Slides.Item(5)
Add-Text $s "三层证据权责" 135 145 760 80 56 $brown $true | Out-Null
Add-Text $s "描述层不可见标签，只负责把可观察事实记录清楚。" 138 245 880 50 28 $muted | Out-Null
Add-Card $s "宏观层" "全片主体、主体属性、整体场景、视觉风格、全局事件、整体音频与敏感内容摘要。" 140 385 500 300
Add-Card $s "片段层" "时间段、动作过程、事件顺序、相邻片段变化、接触关系与音画时间关系。" 710 385 500 300 $true
Add-Card $s "逐帧层" "单帧主体、局部结构、清晰度、画面文字、构图、色彩与相邻帧变化。" 1280 385 500 300
Add-Footer $s 5 "THREE-LAYER EVIDENCE"

# 6 Label ontology
$s = $pres.Slides.Item(6)
Add-Text $s "标签体系：7 个一级标签，31 个二级标签" 760 120 980 80 50 $brown $true | Out-Null
$cats = @(
  "提示词一致性`n主体、属性、动作、场景、风格",
  "视觉生成质量`n清晰度、结构、局部失败、伪影",
  "时间一致性`n角色、物体、场景、动作、事件",
  "物理真实性`n人体、物体、碰撞、自然现象",
  "音频生成质量`n台词、同步、失真、情绪、环境音",
  "审美质量`n构图、镜头、色彩、节奏",
  "安全与合规`n暴力、色情、隐私、版权"
)
for ($i=0; $i -lt $cats.Count; $i++) {
  $col = $i % 2
  $row = [Math]::Floor($i / 2)
  $x = 850 + $col * 455
  $y = 260 + $row * 150
  Add-Box $s $x $y 400 115 $(if($i -eq 6){$brown2}else{$cream}) | Out-Null
  $parts = $cats[$i].Split("`n")
  Add-Text $s $parts[0] ($x+24) ($y+18) 350 36 26 $(if($i -eq 6){$white}else{$brown}) $true | Out-Null
  Add-Text $s $parts[1] ($x+24) ($y+58) 350 38 19 $(if($i -eq 6){$cream}else{$muted}) $false | Out-Null
}
Add-Footer $s 6 "LABEL ONTOLOGY"

# 7 Sample video
$s = $pres.Slides.Item(7)
Add-Text $s "样本视频：真人古装写实短片" 170 95 920 78 52 $brown $true | Out-Null
Add-Box $s 185 220 920 520 $black | Out-Null
Add-Text $s "VEDIO.MP4" 245 390 800 78 58 $white $true "Consolas" 2 | Out-Null
Add-Text $s "视频素材路径：sample\Vedio.mp4" 245 480 800 45 28 $cream $false "Microsoft YaHei UI" 2 | Out-Null
Add-Text $s "为保证模板稳定性，PPT 中保留视频入口说明；原始视频文件与本 PPT 位于同一项目目录。" 245 545 800 70 24 $cream $false "Microsoft YaHei UI" 2 | Out-Null
Add-Box $s 1165 230 540 390 $cream | Out-Null
Add-Text $s "核心提示词要素" 1205 270 420 48 34 $brown $true | Out-Null
Add-BulletList $s @(
  "清晨薄雾中的古代庭院",
  "年轻女书生，浅青交领长衫，木簪束发",
  "竹简出现清晰黑色文字《古风测试》",
  "水缸倒影、竹叶落下、动作连续、画面清晰"
) 1205 340 450 220 23 $black | Out-Null
Add-Footer $s 7 "SAMPLE VIDEO"

# 8 Evidence extraction
$s = $pres.Slides.Item(8)
Add-Text $s "证据抽取：提示词要素核对" 155 140 780 70 52 $brown $true | Out-Null
Add-Text $s "先从 Prompt 提取核心要素，再在三层描述中记录：有、无、不可见、未知。" 158 232 830 46 28 $muted | Out-Null
Add-Card $s "宏观层" "确认古代庭院、服装、竹简等全局要素；指出未出现指定文字《古风测试》。" 150 360 500 270
Add-Card $s "片段层" "按 0-3 秒、3-5 秒、5-8 秒描述动作、场景与道具变化。" 710 360 500 270
Add-Card $s "逐帧层" "帧3：文字模糊不可读，头部被裁切；帧4-5：发型变化为辫子。" 1270 360 500 270 $true
Add-Footer $s 8 "EVIDENCE EXTRACTION"

# 9 Final labels
$s = $pres.Slides.Item(9)
Add-Text $s "裁判输出：最终标签与错因" 118 105 920 76 52 $cream $true | Out-Null
Add-Text $s "最终结果只展示：一级标签--二级标签、证据来源、错因描述。" 120 205 980 46 28 $cream | Out-Null
$tableLeft = 135
$tableTop = 345
$colW = @(410, 330, 800)
$rowH = 105
Add-Box $s $tableLeft $tableTop 1540 64 $brown2 | Out-Null
Add-Text $s "一级标签--二级标签" ($tableLeft+20) ($tableTop+17) 360 35 22 $white $true | Out-Null
Add-Text $s "证据来源" ($tableLeft+$colW[0]+20) ($tableTop+17) 250 35 22 $white $true | Out-Null
Add-Text $s "错因描述" ($tableLeft+$colW[0]+$colW[1]+20) ($tableTop+17) 300 35 22 $white $true | Out-Null
$labels = @(
  @("提示词一致性--属性错误","逐帧层 · 帧3","竹简文字模糊不可读，未识别到指定字样《古风测试》。"),
  @("提示词一致性--动作错误","逐帧层 · 帧3-5","未出现左手按住毛笔动作，且桌面出现非指定剪刀类物件。"),
  @("时间一致性--角色突变","逐帧层 · 帧1-3 vs 帧4-5","人物发型从木簪束发无原因变化为辫子。"),
  @("审美质量--构图问题","逐帧层 · 帧3","人物头部顶部被画面上边缘裁切，影响主体完整性。")
)
for ($i=0; $i -lt $labels.Count; $i++) {
  $y = $tableTop + 64 + $i * $rowH
  Add-Box $s $tableLeft $y 1540 ($rowH-4) $cream | Out-Null
  Add-Text $s $labels[$i][0] ($tableLeft+20) ($y+24) 360 45 22 $brown $true | Out-Null
  Add-Text $s $labels[$i][1] ($tableLeft+$colW[0]+20) ($y+26) 290 44 22 $black | Out-Null
  Add-Text $s $labels[$i][2] ($tableLeft+$colW[0]+$colW[1]+20) ($y+18) 760 62 21 $black | Out-Null
}
Add-Footer $s 9 "JUDGE RESULT"

# 10 Improvements
$s = $pres.Slides.Item(10)
Add-Text $s "关键改进：分批裁判与去重归因" 520 115 980 72 52 $brown $true | Out-Null
Add-Card $s "逐一级标签分批" "每次只关注一个一级标签下的二级标签，降低上下文压力。" 520 285 420 250
Add-Card $s "携带已有打标" "下一批判断当前错因是否已存在，选择替换、保留或移除。" 980 285 420 250 $true
Add-Card $s "紧凑最终输出" "中间过程可分层、分批；最终汇总为一个 JSON 文件。" 1440 285 420 250
Add-Text $s "目标：减少模型反复归因，提高标签结果接近人工标注中的主因判断。" 520 625 980 60 32 $brown2 $true | Out-Null
Add-Footer $s 10 "IMPROVEMENT"

# 11 Limitations
$s = $pres.Slides.Item(11)
Add-Text $s "局限与后续迭代：自适应采样" 890 135 850 76 50 $brown $true | Out-Null
Add-BulletList $s @(
  "逐帧层不是对 24 FPS 视频的全帧穷举",
  "均匀抽帧可能漏掉手部瞬时畸变、物体短暂穿模、文字一帧错误",
  "更可行的策略是《均匀采样 + 变化点采样 + 异常片段加密采样》",
  "当裁判证据不足或低置信时，回到对应时间段补充采样"
) 890 280 820 430 27 $black | Out-Null
Add-Box $s 930 760 720 90 $brown2 | Out-Null
Add-Text $s "优先选择画面骤变帧、不连贯帧和已提示异常的时间段" 970 788 640 42 28 $white $true | Out-Null
Add-Footer $s 11 "LIMITATION"

# 12 Closing
$s = $pres.Slides.Item(12)
Add-Text $s "从视频错误到可追溯证据链" 130 250 1200 130 70 $cream $true | Out-Null
Add-Line $s 134 415 260 $cream | Out-Null
Add-BulletList $s @(
  "描述层负责事实，避免提前带入标签结论",
  "裁判层负责归因，结合 Prompt、标签定义和三层证据",
  "数据看板负责迭代，持续分析准确率、召回率与重复归因"
) 135 485 1050 250 32 $cream | Out-Null
Add-Footer $s 12 "END"

$pres.SaveAs($outPptx, $ppSaveAsOpenXmlPresentation)
$pres.Close()
$pp.Quit()

Write-Output $outPptx













