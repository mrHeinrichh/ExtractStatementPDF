# =====================  MAIN PROGRAM  =====================
# Runs after every function above is defined.
# Test hook: set CF_FOLDER (+ optional CF_MAX) to skip the picker and cap file count.
if ($env:CF_FOLDER) { $SourceFolder = $env:CF_FOLDER } else { $SourceFolder = Pick-Folder }
$MaxFiles = if ($env:CF_MAX) { [int]$env:CF_MAX } else { 0 }
if (-not $SourceFolder) { Write-Host "No folder selected."; return }
if (-not (Test-Path -LiteralPath $SourceFolder)) { Write-Host "Folder not found: $SourceFolder"; return }
$SourceFolder = (Resolve-Path -LiteralPath $SourceFolder).Path

# Output goes to a SEPARATE sibling folder (AR data stays untouched). Override: CF_OUT.
$parent = Split-Path -Parent $SourceFolder; $leaf = Split-Path -Leaf $SourceFolder
$OutputFolder = if ($env:CF_OUT) { $env:CF_OUT } else { Join-Path $parent ($leaf + ' - Frequency') }
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null

try { Get-Contact | Out-Null } catch { Write-Host $_.Exception.Message -ForegroundColor Red; return }
Normalize-Window | Out-Null
Clear-FullTextSearch      # start clean so the Name filter isn't AND-limited by a stale search
Clear-NameFilter

Write-Host "Input (AR data): $SourceFolder"
Write-Host "Output (Excel) : $OutputFolder`n"

$rows = New-Object System.Collections.Generic.List[object]
$xlsFiles = Get-ChildItem -LiteralPath $SourceFolder -Filter *.xls -File | Sort-Object Name
$processed = 0
foreach ($xls in $xlsFiles) {
  try { $cust = Read-XlsCustomer $xls.FullName } catch { $cust = '' }
  if (-not $cust) {
    Write-Host ("SKIP (no name): {0}" -f $xls.Name) -ForegroundColor DarkGray
    $rows.Add([pscustomobject]@{Customer='';Frequency=''}); continue
  }
  try {
    $res = Lookup-Frequency $cust
    if ($res.Freq) {
      $tag = if ($res.Term -eq $cust) { '' } else { " (via '$($res.Term)')" }
      Write-Host ("{0,-40} -> {1}{2}" -f $cust, $res.Freq, $tag) -ForegroundColor Cyan
    } else {
      Write-Host ("{0,-40} -> (not found even after shortening)" -f $cust) -ForegroundColor Yellow
    }
    $rows.Add([pscustomobject]@{Customer=$cust;Frequency=$res.Freq})
    Clear-NameFilter
  } catch {
    Write-Host ("{0} -> ERROR {1}" -f $cust,$_.Exception.Message) -ForegroundColor Red
    $rows.Add([pscustomobject]@{Customer=$cust;Frequency=''})
    try { Clear-NameFilter } catch {}
  }
  $processed++
  if ($MaxFiles -gt 0 -and $processed -ge $MaxFiles) { Write-Host "`n(cap $MaxFiles reached)"; break }
}

# ---- write Excel (.xlsx): two columns, Customer + Frequency ----
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$xlsxPath = Join-Path $OutputFolder ("CustomerFrequency_$stamp.xlsx")
$excel = New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false
try {
  $wb = $excel.Workbooks.Add(); $ws = $wb.Sheets.Item(1); $ws.Name = 'Frequency'
  $headers = @('Customer','Frequency')
  for ($c=0; $c -lt $headers.Count; $c++) { $ws.Cells.Item(1,$c+1) = $headers[$c]; $ws.Cells.Item(1,$c+1).Font.Bold = $true }
  $rIdx = 2
  foreach ($row in $rows) {
    $ws.Cells.Item($rIdx,1) = [string]$row.Customer
    $ws.Cells.Item($rIdx,2) = [string]$row.Frequency
    $rIdx++
  }
  $ws.Columns.Item("A:B").AutoFit() | Out-Null
  $wb.SaveAs($xlsxPath, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
  $wb.Close($true) | Out-Null
} finally { $excel.Quit()|Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null }

# also a CSV copy (same two columns). UTF-8 so names like AVENDAÑO survive.
$csvPath = Join-Path $OutputFolder ("CustomerFrequency_$stamp.csv")
$rows | Select-Object Customer, Frequency | Export-Csv -NoTypeInformation -LiteralPath $csvPath -Encoding UTF8

$found = @($rows | Where-Object { $_.Frequency -ne '' }).Count
$missing = $rows.Count - $found
Write-Host "`n================ DONE ================"
Write-Host ("TOTAL {0}   WITH FREQUENCY {1}   BLANK {2}" -f $rows.Count,$found,$missing)
Write-Host "Excel : $xlsxPath"
Write-Host "CSV   : $csvPath"
