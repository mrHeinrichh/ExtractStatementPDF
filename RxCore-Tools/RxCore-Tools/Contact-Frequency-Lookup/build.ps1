<#
  build.ps1  -  rebuild Contact-Frequency-Lookup.exe from the split source.

  Concatenates every src\*.ps1 (in filename order) into one script
  (Contact-Frequency-Lookup.combined.ps1) and compiles it to the .exe with ps2exe.

  Run:  powershell -ExecutionPolicy Bypass -File build.ps1
#>
$ErrorActionPreference = 'Stop'
$here     = Split-Path -Parent $MyInvocation.MyCommand.Path
$src      = Join-Path $here 'src'
$combined = Join-Path $here 'Contact-Frequency-Lookup.combined.ps1'
$exe      = Join-Path $here 'Contact-Frequency-Lookup.exe'

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
  -title "Contact Frequency Lookup" `
  -description "Look up each customer's Invoice Frequency in Contact.exe and list Name+Frequency to Excel" `
  -company "Plastilens" -product "Contact Frequency Lookup" -version "1.2.0" | Out-Null

if (Test-Path $exe) { Write-Host ("DONE -> {0} ({1:N0} KB)" -f $exe, ((Get-Item $exe).Length/1KB)) }
else { Write-Host "BUILD FAILED" -ForegroundColor Red }
