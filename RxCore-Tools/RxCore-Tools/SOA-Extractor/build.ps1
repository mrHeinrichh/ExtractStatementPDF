<#
  build.ps1  -  rebuild SOA-Extractor.exe from the split source.

  It concatenates every src\*.ps1 (in filename order) into one script
  (SOA-Extractor.combined.ps1) and compiles it to SOA-Extractor.exe with ps2exe.
  The numeric filename prefixes guarantee the correct order (setup, functions, main).

  Run:  powershell -ExecutionPolicy Bypass -File build.ps1
  (Installs the ps2exe module for the current user on first run if needed.)
#>
$ErrorActionPreference = 'Stop'
$here     = Split-Path -Parent $MyInvocation.MyCommand.Path
$src      = Join-Path $here 'src'
$combined = Join-Path $here 'SOA-Extractor.combined.ps1'
$exe      = Join-Path $here 'SOA-Extractor.exe'

Write-Host "Assembling src\*.ps1 ..."
$files = Get-ChildItem -LiteralPath $src -Filter *.ps1 | Sort-Object Name
$sb = New-Object System.Text.StringBuilder
foreach ($f in $files) {
  Write-Host ("  + {0}" -f $f.Name)
  [void]$sb.AppendLine("# ===================== src\$($f.Name) =====================")
  [void]$sb.AppendLine((Get-Content -LiteralPath $f.FullName -Raw))
  [void]$sb.AppendLine('')
}
Set-Content -LiteralPath $combined -Value $sb.ToString() -Encoding UTF8
Write-Host "Wrote $combined"

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
  Write-Host "Installing ps2exe (current user) ..."
  Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
}
Import-Module ps2exe
Invoke-ps2exe -inputFile $combined -outputFile $exe -STA `
  -title "SOA Extractor" `
  -description "Extract Statement of Account CSVs, using each customer's known Invoice Frequency" `
  -company "Plastilens" -product "SOA Extractor" -version "2.0.0" | Out-Null

if (Test-Path $exe) { Write-Host ("DONE -> {0} ({1:N0} KB)" -f $exe, ((Get-Item $exe).Length/1KB)) }
else { Write-Host "BUILD FAILED" -ForegroundColor Red }
