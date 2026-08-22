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

# ConvertFrom-Json gives PSCustomObject, where a missing property silently reads
# as $null -- indistinguishable from an explicit 0 / $false. Specs stay terse by
# omitting knobs, so "was it written?" has to be asked explicitly.
function Test-Prop($obj, [string]$name) {
  return ($null -ne $obj) -and ($obj.PSObject.Properties.Name -contains $name)
}

# Paragraph-level formatting, shared by the body loop and the header/footer filler.
# Every knob is assigned unconditionally for the inheritance reason spelled out at the
# body loop: a paragraph created after another one inherits its formatting, so a
# `if (spec has it) { set it }` shape silently leaks one step of a ladder into the rest.
function Set-ParaFormat($word, $para, $p) {
  $r = $para.Range
  $f = $r.Font
  $f.NameFarEast = $p.fontEA
  $f.NameAscii   = $p.fontLatin
  $f.NameOther   = $p.fontLatin
  $f.Size        = [double]$p.sizePt
  $f.Bold        = [bool]$p.bold

  $fmt = $para.Format
  $fmt.Alignment   = $align[[string]$p.align]
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

$Spec = (Resolve-Path -LiteralPath $Spec).ProviderPath
$outDir = Split-Path -Parent $Output
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
  New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}
$Output = [System.IO.Path]::GetFullPath($Output)

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
  $first = $true
  foreach ($p in $s.paragraphs) {
    if (-not $first) {
      $tailPos = $doc.Content.End - 1
      $tail = $doc.Range($tailPos, $tailPos)
      [void]$tail.InsertParagraphAfter()
    }
    $first = $false
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
