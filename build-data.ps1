<#
  build-data.ps1
  Regenerates data.json (and refreshes the fallback copy embedded in index.html)
  from the KNKT investigation workbook.

  Requirements: Microsoft Excel installed (uses Excel COM automation).

  Usage:
      Right-click -> "Run with PowerShell"
    or from a PowerShell window, in this folder:
      .\build-data.ps1
    optionally point at a workbook elsewhere:
      .\build-data.ps1 -Workbook "D:\path\to\Investigation Database ... .xlsx"

  Workbook layout it expects (sheet 1, header on row 2, data from row 3):
      A Tanggal Peristiwa   B Tahun Peristiwa   E Operator   F Registration
      G Aircraft Type       I Occurrence Type   K Contributing Factors   L Identification
      N Domains             O Disciplines       P Elements
#>

param(
  [string]$Workbook = "$PSScriptRoot\..\Document\Investigation Database - Contributing Factors Update 2026.xlsx",
  [string]$OutDir   = $PSScriptRoot
)

$ErrorActionPreference = "Stop"

function Clean([object]$v) {
  if ($null -eq $v) { return "" }
  return ([string]$v -replace '\s+', ' ').Trim()
}

$wbPath = (Resolve-Path -LiteralPath $Workbook).Path
Write-Host "Reading  $wbPath"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

try {
  $wb = $xl.Workbooks.Open($wbPath, $false, $true)   # read-only
  $s  = $wb.Worksheets.Item(1)

  $lastRow = $s.Cells($s.Rows.Count, 1).End(-4162).Row   # -4162 = xlUp, column A
  if ($lastRow -lt 3) { throw "No data rows found (last row = $lastRow)." }

  # pull the whole block in one shot: columns A..P, rows 3..lastRow
  $vals = $s.Range("A3:P$lastRow").Value2
  $n = $lastRow - 2

  $records = New-Object System.Collections.ArrayList
  for ($i = 1; $i -le $n; $i++) {
    $wbRow = $i + 2

    # date: column A may be an Excel date serial or plain text
    $rawA = $vals.GetValue($i, 1)
    $date = ""
    if ($null -ne $rawA -and "$rawA" -ne "") {
      $num = 0.0
      if ([double]::TryParse([string]$rawA, [ref]$num) -and $num -gt 1) {
        try { $date = [DateTime]::FromOADate($num).ToString("dd MMM yyyy", [System.Globalization.CultureInfo]::InvariantCulture) }
        catch { $date = Clean $rawA }
      } else {
        $date = Clean $rawA
      }
    }

    $rec = [ordered]@{
      row        = $wbRow
      date       = $date
      year       = Clean $vals.GetValue($i, 2)
      operator   = Clean $vals.GetValue($i, 5)
      reg        = Clean $vals.GetValue($i, 6)
      actype     = Clean $vals.GetValue($i, 7)
      occ        = Clean $vals.GetValue($i, 9)
      factor     = Clean $vals.GetValue($i, 11)
      ident      = Clean $vals.GetValue($i, 12)
      domain     = Clean $vals.GetValue($i, 14)
      discipline = Clean $vals.GetValue($i, 15)
      element    = Clean $vals.GetValue($i, 16)
    }
    [void]$records.Add([pscustomobject]$rec)
  }

  $wb.Close($false)
}
finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

# --- write data.json ------------------------------------------------------------
$json = $records | ConvertTo-Json -Compress -Depth 4
if ($records.Count -eq 1) { $json = "[$json]" }   # ConvertTo-Json unwraps single-item arrays

$utf8 = New-Object System.Text.UTF8Encoding($false)
$dataPath = Join-Path $OutDir "data.json"
[System.IO.File]::WriteAllText($dataPath, $json, $utf8)
Write-Host ("Wrote    {0}  ({1} records)" -f $dataPath, $records.Count)

# --- refresh the fallback copy inside index.html ------------------------------
$htmlPath = Join-Path $OutDir "index.html"
if (Test-Path $htmlPath) {
  $lines = [System.IO.File]::ReadAllLines($htmlPath)
  for ($k = 0; $k -lt $lines.Count; $k++) {
    if ($lines[$k] -match '<script id="data" type="application/json">') {
      $lines[$k] = '<script id="data" type="application/json">' + $json + '</script>'
      [System.IO.File]::WriteAllLines($htmlPath, $lines, $utf8)
      Write-Host "Updated  $htmlPath  (embedded fallback copy)"
      break
    }
  }
}

# --- summary -----------------------------------------------------------------
Write-Host ""
Write-Host "Domain distribution:"
$records | Group-Object domain | Sort-Object Count -Descending |
  ForEach-Object { "  {0,-40} {1}" -f $_.Name, $_.Count }
$years = $records | ForEach-Object { $_.year } | Where-Object { $_ } | Sort-Object
if ($years) { Write-Host ("Years: {0} - {1}" -f $years[0], $years[-1]) }
Write-Host ""
Write-Host "Done. Commit data.json + index.html and push to GitHub."
