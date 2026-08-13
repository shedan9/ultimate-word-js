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
$wdLineSpaceSingle   = 0
$wdLayoutModeLineGrid = 2
$wdLayoutModeDefault  = 0
$align = @{ left = 0; center = 1; right = 2; justify = 3; distribute = 4 }

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
  # The Chinese Normal template ships with a line grid switched ON (39 lines,
  # linePitch 312 twips). Leaving LayoutMode untouched means every baseline gets
  # snapped to a 15.6pt multiple, which hides whatever the font metrics say --
  # so a spec without a `grid` must disable it explicitly, not just stay silent.
  if ($s.page.PSObject.Properties.Name -contains 'grid' -and $s.page.grid) {
    $ps.LayoutMode = $wdLayoutModeLineGrid
    $ps.LinesPage  = [int]$s.page.grid.linesPage
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
    $r = $para.Range
    [void]$r.InsertBefore($p.text)
    $para = $doc.Paragraphs.Item($doc.Paragraphs.Count)
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
    if ($p.lineSpacingPt -and [double]$p.lineSpacingPt -gt 0) {
      $fmt.LineSpacingRule = $wdLineSpaceExactly
      $fmt.LineSpacing     = [double]$p.lineSpacingPt
    } else {
      $fmt.LineSpacingRule = $wdLineSpaceSingle
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
