<#
.SYNOPSIS
  Build a .docx fixture from a UTF-8 JSON spec using the installed Word.

.NOTES
  The document text lives in the JSON spec (read as UTF-8 explicitly), never in
  this script, so the fixture content is immune to PowerShell host encoding.
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Spec,
  [Parameter(Mandatory = $true)][string]$Output
)

$ErrorActionPreference = 'Stop'

# --- Word enums -------------------------------------------------------------
$wdAlertsNone        = 0
$wdDoNotSaveChanges  = 0
$wdFormatXMLDocument = 12
$wdLineSpaceExactly  = 4
$wdLineSpaceMultiple = 5
$wdLineSpaceSingle   = 0
$wdLayoutModeLineGrid = 2
$wdLayoutModeGrid     = 3
$wdLayoutModeDefault  = 0
$wdHeaderFooterPrimary   = 1
$wdHeaderFooterFirstPage = 2
$wdHeaderFooterEvenPages = 3
$align = @{ left = 0; center = 1; right = 2; justify = 3; distribute = 4 }
$wdWrapNone          = 3
$msoFalse            = 0
$msoSendBehindText   = 5
# wdRelativeHorizontalPosition / wdRelativeVerticalPosition。两张表的键**不一样**
# （横向有 column / character，纵向有 paragraph / line），照抄一份会静默取到错的框。
$relH = @{ margin = 0; page = 1; column = 2; character = 3; leftMargin = 4; rightMargin = 5; insideMargin = 6; outsideMargin = 7 }
$relV = @{ margin = 0; page = 1; paragraph = 2; line = 3; topMargin = 4; bottomMargin = 5; insideMargin = 6; outsideMargin = 7 }
# --- table enums ------------------------------------------------------------
$wdPreferredWidthPoints = 3
$wdWord9TableBehavior   = 1
$wdAutoFitFixed         = 0
$wdAdjustNone           = 0
$wdRowHeightAuto        = 0
$wdRowHeightAtLeast     = 1
$wdRowHeightExactly     = 2
$wdLineStyleNone        = 0
$wdLineStyleSingle      = 1
# WdLineStyle 的常用几种。名字与 OOXML 的 w:val 对齐（single / dashed / dotted / double…），
# 于是 spec 里写的就是最后落到 w:tcBorders 上的那个词，不必在脑子里翻一次表。
$lineStyleMap = @{
  none = 0; nil = 0; single = 1; dotted = 2; dashSmallGap = 3; dashed = 4
  dashDot = 5; dashDotDot = 6; double = 7; triple = 8; thick = 1
}
# WdBorderType 的码：**全是负数**，顺序是 -1 上 / -2 左 / -3 下 / -4 右（顺时针），
# 之后才是 -5 内横 / -6 内竖。四条边设成同一个值时这张表错没错**看不出来** ——
# 这份文件里原来就写着「-1 左 / -2 右」的注释，直到 spike-table-03 逐边设不同的边框
# 才露馅：要「左格的右边红」，Word 收到的是「左格的**左**边红」，
# 导出的 PDF 里那条线安安静静画在了表格外沿上，看上去像是 Word 不认这条边框。
$cellBorderMap  = @{ top = -1; left = -2; bottom = -3; right = -4 }
$tableBorderMap = @{ top = -1; left = -2; bottom = -3; right = -4; insideH = -5; insideV = -6 }
$wdStyleTypeTable       = 3
$vAlignMap = @{ top = 0; center = 1; bottom = 3 }
# WdConditionCode -> 它落到的 OOXML w:tblStylePr/@w:type。
#
# 这张表**猜不得**，实测的顺序与「按 ECMA 的类型表排」和「按 MSDN 的名字排」**都不一样**：
# 码从 **0** 开始（不是 1），而且 band1Horz / band2Horz 排在 firstCol 之前。
# 探法是把 0..11 各设一个独一无二的字号存盘，解压读 word/styles.xml 认回来（12 及以上、
# 负数一律「参数无效」）。设错一个码，那份格式会安安静静落到**另一个条件**上，
# 样本量出来的数字全对不上，却没有任何输出指得出原因。
$condMap = @{
  firstRow = 0; lastRow = 1; band1Horz = 2; band2Horz = 3
  firstCol = 4; lastCol = 5; band1Vert = 6; band2Vert = 7
  neCell = 8; nwCell = 9; seCell = 10; swCell = 11
}

# ConvertFrom-Json gives PSCustomObject, where a missing property silently reads
# as $null -- indistinguishable from an explicit 0 / $false. Specs stay terse by
# omitting knobs, so "was it written?" has to be asked explicitly.
function Test-Prop($obj, [string]$name) {
  return ($null -ne $obj) -and ($obj.PSObject.Properties.Name -contains $name)
}

# 一条边：样式 + 线宽 + 颜色。
#
# 颜色是这份样本的**读数**：边框冲突的两侧各给一个独一无二的颜色，PDF 里画出来的那条线
# 是什么颜色，就直接说出「赢的是哪一侧」—— 与 spike-table-02 拿字号认条件格式同一招，
# 不必从线宽反推（相邻两格可以配成同宽不同样式，线宽根本分不开）。
#
# Border.Color 的字节序**实测过**：给它 255（照「WdColor 是 BGR」的说法算出来的红）,
# Word 存进 w:color 的是 `0000FF` —— 蓝。也就是这一路上低字节是**蓝**、高字节是红，
# 与 RRGGBB 同序。红蓝写反了看上去毫无破绽（每条结论都读成相反那一侧），
# 所以这里按实测写 r*65536，并留下这行注释，免得下次照文档改回去。
function Set-BorderSide($borders, [int]$code, $spec, $lineStyleMap) {
  $bd = $borders.Item($code)
  $styleName = if ($spec.PSObject.Properties.Name -contains 'style') { [string]$spec.style } else { 'single' }
  if (-not $lineStyleMap.ContainsKey($styleName)) { throw "unknown border style: $styleName" }
  $bd.LineStyle = $lineStyleMap[$styleName]
  if ($bd.LineStyle -eq 0) { return }   # 无边框时线宽与颜色都会被拒
  if ($spec.PSObject.Properties.Name -contains 'widthPt') {
    $bd.LineWidth = [int]([double]$spec.widthPt * 8)
  }
  if ($spec.PSObject.Properties.Name -contains 'color') {
    $hex = [string]$spec.color -replace '^#', ''
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    $bd.Color = $b + ($g * 256) + ($r * 65536)
  }
}

# 一组边（单元格的四条 / 表格的六条），按 spec 里出现的那几条设，没写的**不碰** ——
# 「没写」与「写了 none」在 OOXML 里是两件事（前者退到表级，后者是明确的 w:val="nil"），
# 而这份样本要量的正是这两者的区别。
function Set-Borders($borders, $spec, $map, $lineStyleMap) {
  foreach ($side in $spec.PSObject.Properties.Name) {
    if (-not $map.ContainsKey($side)) { throw "unknown border side: $side" }
    Set-BorderSide $borders $map[$side] $spec.$side $lineStyleMap
  }
}

# Paragraph-level formatting, shared by the body loop and the header/footer filler.
# Every knob is assigned unconditionally for the inheritance reason spelled out at the
# body loop: a paragraph created after another one inherits its formatting, so a
# `if (spec has it) { set it }` shape silently leaks one step of a ladder into the rest.
function Set-ParaFormat($word, $para, $p) {
  $r = $para.Range
  # inheritFont 的块**一个字体属性都不碰**。写上去的话就是直接格式，而直接格式排在
  # 样式链的最后 —— 表格条件格式（w:tblStylePr）会被安静地全部盖掉，spike-table-02
  # 那张「每个条件一个字号」的表会整片显示成同一个字号，看上去像是条件根本没生效。
  if (-not ((Test-Prop $p 'inheritFont') -and [bool]$p.inheritFont)) {
    $f = $r.Font
    $f.NameFarEast = $p.fontEA
    $f.NameAscii   = $p.fontLatin
    $f.NameOther   = $p.fontLatin
    $f.Size        = [double]$p.sizePt
    $f.Bold        = [bool]$p.bold
  }

  $fmt = $para.Format
  # Cells often inherit everything but alignment from the block above; `$align['']`
  # is $null and assigning that throws, so the default is written explicitly rather
  # than made conditional -- a conditional would reintroduce the inheritance leak.
  $fmt.Alignment   = if (Test-Prop $p 'align') { $align[[string]$p.align] } else { 0 }
  $fmt.SpaceBefore = [double]$p.spaceBeforePt
  $fmt.SpaceAfter  = [double]$p.spaceAfterPt
  $fmt.CharacterUnitFirstLineIndent = [double]$p.firstLineChars
  $fmt.RightIndent = [double]$p.rightIndentPt
  $fmt.LeftIndent  = [double]$p.leftIndentPt
  $fmt.PageBreakBefore = [bool]$p.pageBreakBefore
  $fmt.DisableLineHeightGrid = (Test-Prop $p 'snapToGrid') -and (-not [bool]$p.snapToGrid)
  $fmt.WidowControl = -not ((Test-Prop $p 'widowControl') -and (-not [bool]$p.widowControl))
  $fmt.KeepWithNext = [bool]$p.keepWithNext
  $fmt.KeepTogether = [bool]$p.keepTogether
  if ((Test-Prop $p 'lineSpacingMultiple') -and [double]$p.lineSpacingMultiple -gt 0) {
    $fmt.LineSpacingRule = $wdLineSpaceMultiple
    $fmt.LineSpacing     = $word.LinesToPoints([double]$p.lineSpacingMultiple)
  } elseif ($p.lineSpacingPt -and [double]$p.lineSpacingPt -gt 0) {
    $fmt.LineSpacingRule = $wdLineSpaceExactly
    $fmt.LineSpacing     = [double]$p.lineSpacingPt
  } else {
    $fmt.LineSpacingRule = $wdLineSpaceSingle
  }
}

# Fill one header/footer story.
#
# The text of all paragraphs goes in with a single assignment separated by CR: a header
# range starts life with exactly one (empty) paragraph, and InsertParagraphAfter on it
# behaves differently from the body story. One assignment sidesteps that entirely.
#
# A `field` property appends a real Word field ({ PAGE }) at the end of that paragraph --
# typing "1" would not be the same thing at all, and the whole point of these fixtures is
# to see what Word computes for the field on each page.
function Set-StoryContent($word, $hf, $paras) {
  if (-not $paras) { return }
  $texts = @()
  foreach ($p in $paras) { $texts += [string]$p.text }
  $hf.Range.Text = ($texts -join "`r")

  for ($i = 0; $i -lt $paras.Count; $i++) {
    $p = $paras[$i]
    $para = $hf.Range.Paragraphs.Item($i + 1)
    Set-ParaFormat $word $para $p
    if (Test-Prop $p 'field') {
      # End - 1 is just before the paragraph mark; collapsing the paragraph range to its
      # end would instead land at the start of the *next* paragraph.
      $at = $hf.Range.Duplicate
      $at.SetRange($para.Range.End - 1, $para.Range.End - 1)
      [void]$hf.Range.Fields.Add($at, -1, [string]$p.field, $true)
    }
  }
}

# Insert pictures into one paragraph.
#
# `afterChars` means "after the Nth character of the paragraph text", and the pictures are
# inserted BACK TO FRONT: each insertion adds one character, so doing the earlier offsets
# first would silently shift every later one.
#
# Width and height are always written explicitly with LockAspectRatio off first -- the
# pixel size of the file and `wp:extent` are two different things, and with the ratio
# locked Word adjusts the other dimension behind your back, so the sample would no longer
# measure the number the spec asked for.
#
# A `float` block converts it to a floating shape (wdWrapNone = the "in front of / behind
# text" family). **RelativeHorizontalPosition has to be set before Left**: Left is measured
# against whatever frame is currently selected, so the other order makes Word reinterpret
# an already-written coordinate against a new frame and every measured position is wrong.
function Add-ParaImages($doc, $para, $p, $specDir) {
  if (-not (Test-Prop $p 'images')) { return }
  $imgs = @($p.images) | Sort-Object -Property @{ Expression = { [int]$_.afterChars } } -Descending
  foreach ($img in $imgs) {
    $file = [System.IO.Path]::GetFullPath((Join-Path $specDir ([string]$img.file)))
    $start = $para.Range.Start + [int]$img.afterChars
    $at = $doc.Range($start, $start)
    $shape = $doc.InlineShapes.AddPicture($file, $false, $true, $at)
    $shape.LockAspectRatio = $msoFalse
    $shape.Width  = [double]$img.widthPt
    $shape.Height = [double]$img.heightPt
    if (Test-Prop $img 'positionPt') { $shape.Range.Font.Position = [double]$img.positionPt }
    if (Test-Prop $img 'float') {
      $f = $img.float
      $sh = $shape.ConvertToShape()
      $sh.WrapFormat.Type = $wdWrapNone
      $sh.WrapFormat.AllowOverlap = $true
      $sh.LockAspectRatio = $msoFalse
      $sh.RelativeHorizontalPosition = $relH[[string]$f.relativeH]
      $sh.RelativeVerticalPosition   = $relV[[string]$f.relativeV]
      $sh.Left = [double]$f.leftPt
      $sh.Top  = [double]$f.topPt
      $sh.Width  = [double]$img.widthPt
      $sh.Height = [double]$img.heightPt
      if ((Test-Prop $f 'behindDoc') -and [bool]$f.behindDoc) { $sh.ZOrder($msoSendBehindText) }
    }
  }
}

# 填一个单元格：文字（可以是好几段）+ 段落级格式。
#
# 格内的段落靠**一次 Range.Text 赋值里的 CR** 造出来，与 Set-StoryContent 同一招 ——
# 在格子里调 InsertParagraphAfter 会走出格子边界，把**整行**劈成两行，
# 于是一张 4 行的表安安静静变成 5 行，而 spec 看上去毫无问题。
function Set-CellContent($word, $cell, $c) {
  $paras = if (Test-Prop $c 'paras') { @($c.paras) } else { @($c) }
  $texts = @()
  foreach ($p in $paras) { $texts += [string]$p.text }
  # 走 Cell.Range 而不是自己算区间：格子末尾有两个看不见的「单元格结束标记」，
  # 连它们一起覆盖掉的话表结构就坏了。
  $cell.Range.Text = ($texts -join [string][char]13)
  # 新表的格子**继承插入点那一段的直接格式**（表格是插在一个已经排过版的空段落上的），
  # 于是「什么都不写」得到的不是「样式说了算」，而是「上一段的字体与字号」——
  # 它照样是直接格式，照样把表格样式与条件格式盖光。Reset() 把这层刮掉，
  # 让格子真的落回样式链上。inheritFont 的语义就是这个 Reset，不是「少设几个属性」。
  if ((Test-Prop $c 'inheritFont') -and [bool]$c.inheritFont) { $cell.Range.Font.Reset() }
  for ($i = 0; $i -lt $paras.Count; $i++) {
    Set-ParaFormat $word $cell.Range.Paragraphs.Item($i + 1) $paras[$i]
  }
  if (Test-Prop $c 'vAlign') { $cell.VerticalAlignment = $vAlignMap[[string]$c.vAlign] }
  if (Test-Prop $c 'marginPt') {
    $m = $c.marginPt
    if (Test-Prop $m 'top')    { $cell.TopPadding    = [double]$m.top }
    if (Test-Prop $m 'left')   { $cell.LeftPadding   = [double]$m.left }
    if (Test-Prop $m 'bottom') { $cell.BottomPadding = [double]$m.bottom }
    if (Test-Prop $m 'right')  { $cell.RightPadding  = [double]$m.right }
  }
  # 单格的粗边框只为回答一个几何量不出来的问题：**边框吃不吃格内可用宽**。
  # Word 把线画在格线上，所以答案应当是「不吃」—— @uw/layout 一直照这条写，
  # 但 table.ts 的文件头明说它「没有真值」。一格 6pt 的边就能钉死。
  if (Test-Prop $c 'borderWidthPt') {
    foreach ($b in @(-1, -2, -3, -4)) {   # wdBorderTop / Left / Bottom / Right
      $bd = $cell.Borders.Item($b)
      $bd.LineStyle = $wdLineStyleSingle
      $bd.LineWidth = [int]([double]$c.borderWidthPt * 8)
    }
  }
  # 逐边的样式 / 线宽 / 颜色：格线冲突的样本靠它给相邻两格配不同的边
  if (Test-Prop $c 'borders') { Set-Borders $cell.Borders $c.borders $cellBorderMap $lineStyleMap }
}

# 按 spec 造一份自定义表格样式，每个具名条件写成一条 w:tblStylePr。
#
# 条件设的是**字号**而不是加粗：PDF 真值每个片段都带着 size，于是「这一格命中了哪个条件」
# 可以直接从真值上读出来，不必从字形宽度反推 —— 反推会把待标定的度量牵进来。
function Add-TableStyle($doc, $st) {
  $style = $doc.Styles.Add([string]$st.name, $wdStyleTypeTable)
  $t = $style.Table
  if (Test-Prop $st 'rowBandSize') { $t.RowStripe    = [int]$st.rowBandSize }
  if (Test-Prop $st 'colBandSize') { $t.ColumnStripe = [int]$st.colBandSize }
  if (Test-Prop $st 'sizePt')      { $style.Font.Size = [double]$st.sizePt }
  if (Test-Prop $st 'conditions') {
    foreach ($name in $st.conditions.PSObject.Properties.Name) {
      # 0 是合法码（firstRow），所以「没这个条件」只能问 ContainsKey ——
      # 拿 $null -eq $code 判会把 firstRow 一起判成未知条件。
      if (-not $condMap.ContainsKey($name)) { throw "unknown table style condition: $name" }
      $code = $condMap[$name]
      $spec = $st.conditions.$name
      $cond = $t.Condition($code)
      if (Test-Prop $spec 'sizePt') { $cond.Font.Size = [double]$spec.sizePt }
      if (Test-Prop $spec 'bold')   { $cond.Font.Bold = [bool]$spec.bold }
    }
  }
  return [string]$st.name
}

# 在 $at 处插一张表。
#
# 两处顺序不能反：**样式在列宽之前**（表格样式自己可能带列宽，反过来会把 spec 的宽度盖掉）、
# **合并在填字之前**（合并会把两格的文字接起来）。同一行里的合并还要**从右往左**做：
# 合并一次，它右边每一格的下标都要往左挪一个跨度，从左往右做的话第二次就指错格子了。
function Add-Table($word, $doc, $t, $at) {
  $cols = @($t.widthsPt).Count
  $rows = @($t.rows)
  $tbl = $doc.Tables.Add($at, $rows.Count, $cols, $wdWord9TableBehavior, $wdAutoFitFixed)
  $tbl.AllowAutoFit = $false

  if (Test-Prop $t 'style') { $tbl.Style = Add-TableStyle $doc $t.style }
  if (Test-Prop $t 'look') {
    $lk = $t.look
    $tbl.ApplyStyleHeadingRows = [bool]$lk.firstRow
    $tbl.ApplyStyleLastRow     = [bool]$lk.lastRow
    $tbl.ApplyStyleFirstColumn = [bool]$lk.firstColumn
    $tbl.ApplyStyleLastColumn  = [bool]$lk.lastColumn
    $tbl.ApplyStyleRowBands    = [bool]$lk.rowBands
    $tbl.ApplyStyleColumnBands = [bool]$lk.colBands
  }

  # PreferredWidth 也得用点：留着默认的百分比，Word 会照正文栏宽反推列宽，
  # 这份样本要钉死的那串 w:tblGrid 就成了 Word 随手给的数。
  $total = 0.0
  foreach ($w in $t.widthsPt) { $total += [double]$w }
  $tbl.PreferredWidthType = $wdPreferredWidthPoints
  $tbl.PreferredWidth = $total
  for ($i = 0; $i -lt $cols; $i++) {
    $tbl.Columns.Item($i + 1).SetWidth([double]$t.widthsPt[$i], $wdAdjustNone)
  }

  if (Test-Prop $t 'align')    { $tbl.Rows.Alignment  = $align[[string]$t.align] }
  if (Test-Prop $t 'indentPt') { $tbl.Rows.LeftIndent = [double]$t.indentPt }
  if (Test-Prop $t 'cellMarginPt') {
    $m = $t.cellMarginPt
    if (Test-Prop $m 'top')    { $tbl.TopPadding    = [double]$m.top }
    if (Test-Prop $m 'left')   { $tbl.LeftPadding   = [double]$m.left }
    if (Test-Prop $m 'bottom') { $tbl.BottomPadding = [double]$m.bottom }
    if (Test-Prop $m 'right')  { $tbl.RightPadding  = [double]$m.right }
  }
  if (Test-Prop $t 'borderWidthPt') {
    $w = [int]([double]$t.borderWidthPt * 8)
    $tbl.Borders.InsideLineStyle  = $wdLineStyleSingle
    $tbl.Borders.OutsideLineStyle = $wdLineStyleSingle
    $tbl.Borders.InsideLineWidth  = $w
    $tbl.Borders.OutsideLineWidth = $w
  }
  # 表级逐边（含 insideH / insideV）。放在 borderWidthPt 之后：两者同时写时以细的这份为准
  if (Test-Prop $t 'borders') { Set-Borders $tbl.Borders $t.borders $tableBorderMap $lineStyleMap }

  for ($ri = 0; $ri -lt $rows.Count; $ri++) {
    $r = $rows[$ri]
    $cells = @($r.cells)
    # 每个 spec 格子起始的**网格列号**：于是 spec 只说「这一格跨几列」，
    # 不必去数 Word 那套一合并就变的下标。
    $starts = @()
    $g = 1
    foreach ($c in $cells) {
      $starts += $g
      if (Test-Prop $c 'span') { $g += [int]$c.span } else { $g += 1 }
    }
    for ($k = $cells.Count - 1; $k -ge 0; $k--) {
      $span = 1
      if (Test-Prop $cells[$k] 'span') { $span = [int]$cells[$k].span }
      if ($span -gt 1) {
        $s0 = $starts[$k]
        $tbl.Cell($ri + 1, $s0).Merge($tbl.Cell($ri + 1, $s0 + $span - 1))
      }
    }

    $row = $tbl.Rows.Item($ri + 1)
    if (Test-Prop $r 'header')    { $row.HeadingFormat = [bool]$r.header }
    if (Test-Prop $r 'cantSplit') { $row.AllowBreakAcrossPages = -not [bool]$r.cantSplit }
    if (Test-Prop $r 'heightPt') {
      $rule = 'atLeast'
      if (Test-Prop $r 'heightRule') { $rule = [string]$r.heightRule }
      if ($rule -eq 'exact') { $row.HeightRule = $wdRowHeightExactly }
      else                   { $row.HeightRule = $wdRowHeightAtLeast }
      $row.Height = [double]$r.heightPt
    } else {
      $row.HeightRule = $wdRowHeightAuto
    }
    for ($k = 0; $k -lt $cells.Count; $k++) {
      Set-CellContent $word $row.Cells.Item($k + 1) $cells[$k]
    }
  }
}

$Spec = (Resolve-Path -LiteralPath $Spec).ProviderPath
$outDir = Split-Path -Parent $Output
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
  New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}
$Output = [System.IO.Path]::GetFullPath($Output)

$specDir = Split-Path -Parent $Spec
$json = [System.IO.File]::ReadAllText($Spec, [System.Text.Encoding]::UTF8)
$s = $json | ConvertFrom-Json

$word = $null
$doc  = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = $wdAlertsNone
  $doc = $word.Documents.Add()

  # --- page setup -----------------------------------------------------------
  $ps = $doc.PageSetup
  $mm = { param($v) $word.MillimetersToPoints($v) }
  $ps.PageWidth    = & $mm $s.page.widthMm
  $ps.PageHeight   = & $mm $s.page.heightMm
  $ps.TopMargin    = & $mm $s.page.marginMm.top
  $ps.BottomMargin = & $mm $s.page.marginMm.bottom
  $ps.LeftMargin   = & $mm $s.page.marginMm.left
  $ps.RightMargin  = & $mm $s.page.marginMm.right
  # w:pgMar/@w:header is the distance from the paper edge to the TOP of the header,
  # @w:footer the distance to the BOTTOM of the footer -- the two are not symmetric,
  # which is exactly one of the things spike-header-01 is here to prove.
  if (Test-Prop $s.page 'headerDistMm') { $ps.HeaderDistance = & $mm $s.page.headerDistMm }
  if (Test-Prop $s.page 'footerDistMm') { $ps.FooterDistance = & $mm $s.page.footerDistMm }
  $ps.DifferentFirstPageHeaderFooter = [bool]$s.page.differentFirstPage
  $ps.OddAndEvenPagesHeaderFooter    = [bool]$s.page.differentOddEven
  # The Chinese Normal template ships with a line grid switched ON (39 lines,
  # linePitch 312 twips). Leaving LayoutMode untouched means every baseline gets
  # snapped to a 15.6pt multiple, which hides whatever the font metrics say --
  # so a spec without a `grid` must disable it explicitly, not just stay silent.
  if ((Test-Prop $s.page 'grid') -and $s.page.grid) {
    $g = $s.page.grid
    # "22 lines x 28 chars" needs the *character* grid (wdLayoutModeGrid), which
    # snaps glyphs to columns as well; a line count alone only snaps baselines.
    # In a real gongwen both are on, and they are separate calibration questions --
    # so the spec picks one, it is not inferred.
    $ps.LinesPage = [int]$g.linesPage
    if ((Test-Prop $g 'charsLine') -and $g.charsLine) {
      $ps.LayoutMode = $wdLayoutModeGrid
      $ps.CharsLine  = [int]$g.charsLine
    } else {
      $ps.LayoutMode = $wdLayoutModeLineGrid
    }
  } else {
    $ps.LayoutMode = $wdLayoutModeDefault
  }

  # --- paragraphs -----------------------------------------------------------
  # Always append an empty paragraph at the very end and fill *that* one.
  # Paragraphs.Add() without a Range anchors on the selection, which is still at
  # offset 0 right after Documents.Add() -- that silently merges paragraph 2 into
  # paragraph 1. Being explicit about the tail position avoids the whole class of bug.
  # 带 "kind": "table" 的块是一张表，其余一律当段落 —— 于是表格进来之前写的
  # 十八份 spec 一个字都不用改。
  $needPara = $false
  foreach ($p in $s.paragraphs) {
    if ($needPara) {
      $tailPos = $doc.Content.End - 1
      $tail = $doc.Range($tailPos, $tailPos)
      [void]$tail.InsertParagraphAfter()
    }
    $needPara = $true

    if ((Test-Prop $p 'kind') -and ([string]$p.kind -eq 'table')) {
      $tailPos = $doc.Content.End - 1
      Add-Table $word $doc $p $doc.Range($tailPos, $tailPos)
      # 表后面 Word 一定留一个空段落（文档不能以表结尾），下一块直接用它。
      # 再插一个的话每张表后面都平白多一个空行 —— 开着行网格的页面上
      # 那就是整整一个网格行的版心。
      $needPara = $false
      continue
    }

    $para = $doc.Paragraphs.Item($doc.Paragraphs.Count)

    # NB: $r.Text = ... replaces the paragraph mark too, collapsing every
    # paragraph into one. InsertBefore keeps the paragraph structure intact.
    # Empty text is legal and load-bearing: an empty paragraph's height comes
    # entirely from the paragraph mark's own character properties, and the font
    # assignment below is how that mark gets them.
    $r = $para.Range
    if ($p.text) { [void]$r.InsertBefore($p.text) }
    $para = $doc.Paragraphs.Item($doc.Paragraphs.Count)
    $r = $para.Range

    # Formatting is applied by the shared helper -- the body loop and the header/footer
    # filler must not drift apart, or a fixture's header would silently get a different
    # line spacing rule from its body and the geometry it measures would mean nothing.
    Set-ParaFormat $word $para $p

    # Pictures go in after the formatting: the paragraph's font size defines the line box and
    # a picture merely drops a taller box into it. The other order would let Set-ParaFormat
    # spread the run properties over the picture's own character as well.
    Add-ParaImages $doc $para $p $specDir
  }

  # --- headers / footers ----------------------------------------------------
  # Written AFTER the body so the page count is already what it will be: assigning to a
  # header's Range repaginates, and Word only materialises the first-page / even-page
  # stories once the corresponding PageSetup switch is on (done above).
  $sec = $doc.Sections.Item(1)
  $stories = @(
    @{ set = $s.headers; coll = $sec.Headers },
    @{ set = $s.footers; coll = $sec.Footers }
  )
  $types = @{ 'default' = $wdHeaderFooterPrimary; 'first' = $wdHeaderFooterFirstPage; 'even' = $wdHeaderFooterEvenPages }
  foreach ($story in $stories) {
    if (-not $story.set) { continue }
    foreach ($kind in @('default', 'first', 'even')) {
      if (-not (Test-Prop $story.set $kind)) { continue }
      Set-StoryContent $word $story.coll.Item($types[$kind]) $story.set.$kind
    }
  }

  $doc.SaveAs2($Output, $wdFormatXMLDocument)
  Write-Output "OK $Output"
}
finally {
  if ($doc)  { try { $doc.Close($wdDoNotSaveChanges) } catch {} }
  if ($word) { try { $word.Quit($wdDoNotSaveChanges) } catch {} }
  foreach ($o in @($doc, $word)) {
    if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
