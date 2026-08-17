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
$align = @{ left = 0; center = 1; right = 2; justify = 3; distribute = 4 }

# ConvertFrom-Json gives PSCustomObject, where a missing property silently reads
# as $null -- indistinguishable from an explicit 0 / $false. Specs stay terse by
# omitting knobs, so "was it written?" has to be asked explicitly.
function Test-Prop($obj, [string]$name) {
  return ($null -ne $obj) -and ($obj.PSObject.Properties.Name -contains $name)
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
    # Every paragraph-level knob below is assigned UNCONDITIONALLY, even when the spec
    # omits it. InsertParagraphAfter inherits the previous paragraph's formatting, so a
    # `if (spec has it) { set it }` shape lets one paragraph's setting leak into every
    # later one -- pageBreakBefore on paragraph 4 silently put all 17 on their own page.
    # A missing JSON property reads as $null, and [bool]$null is $false, which is the
    # intended default in both cases.
    #
    # A hard page break makes this paragraph the *first* on its page, so its first
    # baseline is measured from the top margin and nothing else -- that is the whole
    # point of the baseline spike. A break character in the text would instead put
    # the break inside this paragraph, which is not the same thing.
    $fmt.PageBreakBefore = [bool]$p.pageBreakBefore
    # w:snapToGrid lives here under an inverted name: DisableLineHeightGrid = True means
    # 'do not snap'. Needed to prove the grid fixture actually measures snapping and not
    # some coincidence of the font metrics -- one paragraph opts out and must move.
    # Absent property => snap (the Word default), hence the double negation.
    $fmt.DisableLineHeightGrid = (Test-Prop $p 'snapToGrid') -and (-not [bool]$p.snapToGrid)
    if ((Test-Prop $p 'lineSpacingMultiple') -and [double]$p.lineSpacingMultiple -gt 0) {
      # LineSpacing for the multiple rule is in points, hence LinesToPoints --
      # assigning 1.5 directly would mean "1.5pt fixed" and silently collapse the line.
      $fmt.LineSpacingRule = $wdLineSpaceMultiple
      $fmt.LineSpacing     = $word.LinesToPoints([double]$p.lineSpacingMultiple)
    } elseif ($p.lineSpacingPt -and [double]$p.lineSpacingPt -gt 0) {
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
