<#
.SYNOPSIS
  Export a .docx to PDF with the locally installed Word, plus a sidecar JSON of
  Word's own layout facts (page count, page setup) used as truth metadata.

.NOTES
  Text must stay text in the PDF, so BitmapMissingFonts is $false.
  Run from anywhere; paths are resolved to absolute before touching COM.
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$OutputPdf,
  [string]$MetaJson
)

$ErrorActionPreference = 'Stop'

# --- constants (Word enums) -------------------------------------------------
$wdExportFormatPDF      = 17
$wdExportOptimizeForPrint = 0
$wdExportAllDocument    = 0
$wdExportDocumentContent = 0   # no markup
$wdDoNotSaveChanges     = 0
$wdAlertsNone           = 0
$wdStatisticPages       = 2
$wdStatisticWords       = 0
$wdStatisticCharacters  = 3

# --- resolve paths ----------------------------------------------------------
$InputPath = (Resolve-Path -LiteralPath $InputPath).ProviderPath
$outDir = Split-Path -Parent $OutputPdf
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
  New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}
$OutputPdf = [System.IO.Path]::GetFullPath($OutputPdf)
if ($MetaJson) { $MetaJson = [System.IO.Path]::GetFullPath($MetaJson) }

$word = $null
$doc  = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = $wdAlertsNone
  $word.Options.Pagination = $true

  $doc = $word.Documents.Open(
    $InputPath,       # FileName
    $false,           # ConfirmConversions
    $true,            # ReadOnly
    $false,           # AddToRecentFiles
    '', '',           # PasswordDocument, PasswordTemplate
    $false,           # Revert
    '', '',           # WritePasswordDocument, WritePasswordTemplate
    0,                # Format
    0,                # Encoding
    $false            # Visible
  )
  $doc.Repaginate()

  $doc.ExportAsFixedFormat(
    $OutputPdf,
    $wdExportFormatPDF,
    $false,                     # OpenAfterExport
    $wdExportOptimizeForPrint,
    $wdExportAllDocument,
    1, 1,                       # From, To (ignored for AllDocument)
    $wdExportDocumentContent,
    $false,                     # IncludeDocProps
    $true,                      # KeepIRM
    0,                          # CreateBookmarks: none
    $false,                     # DocStructureTags
    $false,                     # BitmapMissingFonts  <- keep glyphs as text
    $false                      # UseISO19005_1 (PDF/A)
  )

  if ($MetaJson) {
    $sections = @()
    foreach ($s in $doc.Sections) {
      $ps = $s.PageSetup
      $sections += [ordered]@{
        pageWidth    = [math]::Round($ps.PageWidth, 4)
        pageHeight   = [math]::Round($ps.PageHeight, 4)
        topMargin    = [math]::Round($ps.TopMargin, 4)
        bottomMargin = [math]::Round($ps.BottomMargin, 4)
        leftMargin   = [math]::Round($ps.LeftMargin, 4)
        rightMargin  = [math]::Round($ps.RightMargin, 4)
        headerDist   = [math]::Round($ps.HeaderDistance, 4)
        footerDist   = [math]::Round($ps.FooterDistance, 4)
        gutter       = [math]::Round($ps.Gutter, 4)
        orientation  = [int]$ps.Orientation
        lineNumbers  = [bool]$ps.LineNumbering.Active
      }
    }
    $meta = [ordered]@{
      source      = [System.IO.Path]::GetFileName($InputPath)
      wordVersion = $word.Version
      wordBuild   = $word.Build
      unit        = 'pt'
      pageCount   = [int]$doc.ComputeStatistics($wdStatisticPages)
      wordCount   = [int]$doc.ComputeStatistics($wdStatisticWords)
      charCount   = [int]$doc.ComputeStatistics($wdStatisticCharacters)
      sections    = $sections
    }
    $json = $meta | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($MetaJson, $json, (New-Object System.Text.UTF8Encoding($false)))
  }

  Write-Output "OK $OutputPdf"
}
finally {
  if ($doc)  { try { $doc.Close($wdDoNotSaveChanges) } catch {} }
  if ($word) { try { $word.Quit($wdDoNotSaveChanges) } catch {} }
  foreach ($o in @($doc, $word)) {
    if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
