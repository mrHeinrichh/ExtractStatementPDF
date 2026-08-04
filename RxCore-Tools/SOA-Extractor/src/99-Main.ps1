# =====================  MAIN PROGRAM  =====================
# Runs after every function above is defined.
# Test hooks (optional): SOA_FOLDER, SOA_MAX, SOA_REF (reference CSV), SOA_OUT.

# ---- 1) Choose the AR folder (the .xls files) ----
if ($env:SOA_FOLDER) { $SourceFolder = $env:SOA_FOLDER } else { $SourceFolder = Pick-Folder }
$MaxFiles = if ($env:SOA_MAX) { [int]$env:SOA_MAX } else { 0 }
if (-not $SourceFolder) { Write-Host "No folder selected. Exiting."; return }
if (-not (Test-Path -LiteralPath $SourceFolder)) { Write-Host "Folder not found: $SourceFolder"; return }
$SourceFolder = (Resolve-Path -LiteralPath $SourceFolder).Path

# ---- 2) Must be logged in to Accounting ----
if (-not (Assert-AcctLoggedIn)) { Write-Host "Not logged in to Accounting. Exiting." -ForegroundColor Red; return }

# ---- 3) Output folder (separate from the AR data). Override: SOA_OUT ----
if ($env:SOA_OUT) {
  $OutputFolder = $env:SOA_OUT
} else {
  $parent = Split-Path -Parent $SourceFolder
  $leaf   = Split-Path -Leaf   $SourceFolder
  $OutputFolder = Join-Path $parent ($leaf + ' - SOA CSV')
}
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
New-Item -ItemType Directory -Force -Path $ShotDir      | Out-Null

# ---- 4) Reference CSV for frequencies (the file exported from Contact) ----
if     ($env:SOA_REF)     { $refPath = $env:SOA_REF }
elseif ($env:SOA_FREQCSV) { $refPath = $env:SOA_FREQCSV }
else   { $refPath = Browse-File "Select the Customer CSV exported from Contact (frequency reference)" }
$ref     = Load-FreqTable $refPath
$FreqMap = $ref.Map

# ---- 5) Load the persistent run log (for resume) ----
$logCsv  = Join-Path $OutputFolder '_SOA_Log.csv'
$logXlsx = Join-Path $OutputFolder '_SOA_Log.xlsx'
$log     = Load-Log $logCsv
$alreadyDone = @($log.Values | Where-Object { $_.Status -eq 'DONE' }).Count

Write-Host "Input (AR data)  : $SourceFolder"
Write-Host "Output (CSV)     : $OutputFolder"
Write-Host "Frequency ref    : $($ref.Source)  ($($FreqMap.Count) names)"
Write-Host "Log (resume from): $logCsv  ($alreadyDone already DONE)"
Write-Host "Fallback order   : $($FrequencyOrder -join ' -> ')`n"

# ---- 6) Process each .xls (skipping ones already extracted) ----
$xlsFiles = Get-ChildItem -LiteralPath $SourceFolder -Filter *.xls -File | Sort-Object Name
$stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
$did = 0; $processed = 0
foreach ($xls in $xlsFiles) {
  $base = [System.IO.Path]::GetFileNameWithoutExtension($xls.Name)
  if ($base.ToLower().EndsWith('.csv')) { $base = $base.Substring(0, $base.Length-4) }
  $csvPath = Join-Path $OutputFolder ($base + '.csv')

  $prior = $log[$xls.Name]
  $isDone = ($prior -and $prior.Status -eq 'DONE') -or (Test-Path -LiteralPath $csvPath)
  if ($isDone -and -not $Overwrite) {
    Write-Host ("SKIP (already extracted): {0}" -f $xls.Name) -ForegroundColor DarkGray
    if (-not $prior) {
      $log[$xls.Name] = [pscustomobject]@{ Customer=$base; File=$xls.Name; DateRange=''; Frequency=''
        Status='DONE'; Reason='CSV already present in output folder'; Csv=(Split-Path $csvPath -Leaf); UpdatedOn=$stamp }
    }
    continue
  }

  try { $meta = Read-XlsMeta $xls.FullName }
  catch {
    Write-Host ("META FAIL {0}: {1}" -f $xls.Name,$_.Exception.Message) -ForegroundColor Red
    $log[$xls.Name] = [pscustomobject]@{ Customer=''; File=$xls.Name; DateRange=''; Frequency=''
      Status='ERROR'; Reason=(Explain 'meta-error' '')[1]; Csv=''; UpdatedOn=$stamp }
    continue
  }

  $range = "{0}..{1}" -f $meta.From, $meta.To
  $r = Resolve-Frequency $meta.Customer $FreqMap
  $known = $r.Freq
  if ($known) { $tryOrder = @($known) + ($FrequencyOrder | Where-Object { $_ -ne $known }) }
  else        { $tryOrder = $FrequencyOrder }

  $freqTag = if ($known) { "freq=$known ($($r.Via))" } else { "freq=unknown -> will guess" }
  Write-Host ("PROCESS: {0}  [{1}]  {2}  {3}" -f $xls.Name,$meta.Customer,$range,$freqTag) -ForegroundColor Cyan

  $status='no-data'; $usedFreq=''
  foreach ($freq in $tryOrder) {
    try {
      Reset-State
      Open-StatementDialog | Out-Null
      Set-Customer $meta.Customer
      if ($freq -ne 'Monthly') { Set-Frequency $freq }
      Set-DateField 'From' $meta.From; Set-DateField 'To' $meta.To
      Shot ("{0}_{1}_form" -f $base,$freq); Click-OK; Start-Sleep -Seconds 5
      if (-not (Get-ReportWindow)) { Write-Host ("   {0}: no report -> next" -f $freq) -ForegroundColor Yellow; Close-Dialog; continue }
      Shot ("{0}_{1}_report" -f $base,$freq)
      $status = Export-Csv-FromReport $csvPath; $usedFreq = $freq; Close-Report
      Write-Host ("   {0}: {1}" -f $freq,$status) -ForegroundColor Green
      if ($status -in @('saved','overwritten')) { break }
    } catch {
      Write-Host ("   {0}: ERROR {1}" -f $freq,$_.Exception.Message) -ForegroundColor Red
      Close-Dialog; Close-Report; $status = "error: $($_.Exception.Message)"
    }
  }

  $ex = Explain $status $usedFreq          # -> @(SUCCESS|FAILED|SKIPPED, reason)
  $logStatus = switch ($ex[0]) { 'SUCCESS' {'DONE'} 'SKIPPED' {'SKIPPED'} default {'ERROR'} }
  $reason = if ($known -and $logStatus -eq 'DONE') { "$($ex[1]); frequency from reference ($($r.Via))" } else { $ex[1] }
  $log[$xls.Name] = [pscustomobject]@{ Customer=$meta.Customer; File=$xls.Name; DateRange=$range; Frequency=$usedFreq
    Status=$logStatus; Reason=$reason; Csv=$(if ($logStatus -eq 'DONE') { Split-Path $csvPath -Leaf } else { '' }); UpdatedOn=$stamp }

  $did++; $processed++
  # Save the log after every file so a crash/close still leaves an up-to-date, resumable log.
  Save-Log ($log.Values) $logCsv $logXlsx
  if ($MaxFiles -gt 0 -and $processed -ge $MaxFiles) { Write-Host "`n(Test cap reached: $MaxFiles files)"; break }
}

# ---- 7) Final save + summary ----
Save-Log ($log.Values) $logCsv $logXlsx
$done  = @($log.Values | Where-Object { $_.Status -eq 'DONE'  }).Count
$err   = @($log.Values | Where-Object { $_.Status -eq 'ERROR' }).Count
$skip  = @($log.Values | Where-Object { $_.Status -eq 'SKIPPED' }).Count

Write-Host "`n================ SUMMARY ================"
Write-Host ("Processed this run : {0}" -f $did)
Write-Host ("DONE (total)       : {0}" -f $done)
Write-Host ("ERROR (need retry) : {0}" -f $err)
Write-Host ("SKIPPED            : {0}" -f $skip)
Write-Host "`nCSV output folder  : $OutputFolder"
Write-Host "Log (Excel)        : $logXlsx"
Write-Host "Log (CSV)          : $logCsv"
Write-Host "Step screenshots   : $ShotDir"
if ($err -gt 0) { Write-Host "`nRe-run to retry the ERROR rows; DONE rows are skipped automatically." -ForegroundColor Yellow }
