<#
  build.ps1  -  rebuild Customer-Matcher.exe from the split source.
  Concatenates src\*.ps1 (filename order) into Customer-Matcher.combined.ps1 and
  compiles it with ps2exe.
  Run:  powershell -ExecutionPolicy Bypass -File build.ps1
#>
$ErrorActionPreference = 'Stop'
$here     = Split-Path -Parent $MyInvocation.MyCommand.Path
$src      = Join-Path $here 'src'
$combined = Join-Path $here 'Customer-Matcher.combined.ps1'
$exe      = Join-Path $here 'Customer-Matcher.exe'

Write-Host "Assembling src\*.ps1 ..."
$sb = New-Object System.Text.StringBuilder
foreach ($f in (Get-ChildItem -LiteralPath $src -Filter *.ps1 | Sort-Object Name)) {
  Write-Host ("  + {0}" -f $f.Name)
  [void]$sb.AppendLine("# ===================== src\$($f.Name) =====================")
  [void]$sb.AppendLine((Get-Content -LiteralPath $f.FullName -Raw))
  [void]$sb.AppendLine('')
}
Set-Content -LiteralPath $combined -Value $sb.ToString() -Encoding UTF8

if (-not (Get-Module -ListAvailable -Name ps2exe)) { Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber }
Import-Module ps2exe
Invoke-ps2exe -inputFile $combined -outputFile $exe -STA `
  -title "Customer Matcher" `
  -description "Annotate the Contact Customer export with the matching SOA .xls names" `
  -company "Plastilens" -product "Customer Matcher" -version "1.0.0" | Out-Null

if (Test-Path $exe) { Write-Host ("DONE -> {0} ({1:N0} KB)" -f $exe, ((Get-Item $exe).Length/1KB)) }
else { Write-Host "BUILD FAILED" -ForegroundColor Red }
