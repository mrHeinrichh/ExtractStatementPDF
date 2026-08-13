# =====================  MAIN PROGRAM  =====================
# Runs after every function above is defined.
# Test hooks (optional): SOA_FOLDER, SOA_MAX, SOA_REF (reference CSV), SOA_OUT.

# ---- 1) Choose the AR folder (the .xls files) ----
if ($env:SOA_FOLDER) { $SourceFolder = $env:SOA_FOLDER } else { $SourceFolder = Pick-Folder }
$MaxFiles = if ($env:SOA_MAX) { [int]$env:SOA_MAX } else { 0 }
if (-not $SourceFolder) { Write-Host "No folder selected. Exiting."; return }
if (-not (Test-Path -LiteralPath $SourceFolder)) { Write-Host "Folder not found: $SourceFolder"; return }
$SourceFolder = (Resolve-Path -LiteralPath $SourceFolder).Path

# ---- 2) Choose which Accounting window to drive, then confirm it's logged in ----
$acct = Choose-Acct
if (-not $acct) {
  if (@(Get-AcctCandidates).Count -eq 0) {
    Show-Message "Accounting is not open.`n`nPlease open Accounting.exe and log in (user hfabros), then run this again."
  }
  Write-Host "No Accounting window selected. Exiting." -ForegroundColor Red; return
}
Write-Host ("Accounting        : PID {0}  {1}" -f $acct.Pid, $acct.Title)

# Elevation/UAC mismatch: can't automate an elevated app from a normal one.
if (Test-ElevationMismatch) {
  Write-Host "`nELEVATION MISMATCH:`n$ElevationHelp" -ForegroundColor Red
  Show-Message $ElevationHelp
  return
}

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

# ---- 4) Reference CSV (RxOffice/Contact export): gives the customer NAME to type
#         in Accounting AND the frequency, matched from each .xls customer name. ----
if     ($env:SOA_REF)     { $refPath = $env:SOA_REF }
elseif ($env:SOA_FREQCSV) { $refPath = $env:SOA_FREQCSV }
else   { $refPath = Browse-File "Select the reference CSV (RxOffice export: Name + StatementFrequency)" }
$ref = Load-Reference $refPath

# ---- 5) Load the persistent run log (for resume) ----
$logCsv  = Join-Path $OutputFolder '_SOA_Log.csv'
$logXlsx = Join-Path $OutputFolder '_SOA_Log.xlsx'
$log     = Load-Log $logCsv
$alreadyDone = @($log.Values | Where-Object { $_.Status -eq 'DONE' }).Count

Write-Host "Input (AR data)  : $SourceFolder"
Write-Host "Output (CSV)     : $OutputFolder"
Write-Host "Reference        : $($ref.Source)  ($($ref.ByKey.Count) names)"
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
      $log[$xls.Name] = [pscustomobject]@{ Customer=$base; RxOfficeName=''; File=$xls.Name; DateRange=''; Frequency=''
        Status='DONE'; Reason='CSV already present in output folder'; Csv=(Split-Path $csvPath -Leaf); UpdatedOn=$stamp }
    }
    continue
  }

  try { $meta = Read-XlsMeta $xls.FullName }
  catch {
    Write-Host ("META FAIL {0}: {1}" -f $xls.Name,$_.Exception.Message) -ForegroundColor Red
    $log[$xls.Name] = [pscustomobject]@{ Customer=''; RxOfficeName=''; File=$xls.Name; DateRange=''; Frequency=''
      Status='ERROR'; Reason=(Explain 'meta-error' '')[1]; Csv=''; UpdatedOn=$stamp }
    continue
  }

  $range = "{0}..{1}" -f $meta.From, $meta.To

  # Match the .xls (AR) customer to the reference: get the RxOffice NAME to type
  # in Accounting and the frequency. Fall back to the raw .xls name if unmatched.
  $r = Resolve-Reference $meta.Customer $ref
  $known   = $r.Freq
  # Type the matched RxOffice name (exact / normalized / word-prefix). Matching is
  # precise (>=2-word prefix), so this is the canonical name Accounting expects
  # (e.g. AR "ABALOS GUILLERMO OPTICAL" -> types "ABALOS GUILLERMO"). Only when there
  # is NO reference match do we fall back to typing the raw .xls name.
  $custName = if ($r.Name) { $r.Name } else { $meta.Customer }
  # Prefer searching by id (2nd/"Code" column of the reference CSV); fall back to the
  # resolved RxOffice name, then the raw .xls name, when no id is available.
  $searchInput = if ($r.Id) { $r.Id } elseif ($r.Name) { $r.Name } else { $meta.Customer }
  if ($known) { $tryOrder = @($known) + ($FrequencyOrder | Where-Object { $_ -ne $known }) }
  else        { $tryOrder = $FrequencyOrder }

  $nameTag = if ($r.Via -eq 'exact' -or $r.Via -eq 'normalized') { " -> ref '$custName'" } elseif ($r.Name) { " -> $($r.Via)" } else { " (no ref match; raw name)" }
  $freqTag = if ($known) { "freq=$known" } else { "freq=unknown -> guess" }
  Write-Host ("PROCESS: {0}  [{1}]{2}  {3}  {4}" -f $xls.Name,$meta.Customer,$nameTag,$range,$freqTag) -ForegroundColor Cyan

  $status='no-data'; $usedFreq=''
  foreach ($freq in $tryOrder) {
    try {
      Reset-State
      Open-StatementDialog | Out-Null
      Set-Customer $searchInput       # <-- type the id from the reference (falls back to name)
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
  $reason = if ($logStatus -eq 'DONE' -and $r.Name) { "$($ex[1]); typed '$searchInput' (match: $($r.Via))" }
            else { $ex[1] }
  $log[$xls.Name] = [pscustomobject]@{ Customer=$meta.Customer; RxOfficeName=$custName; File=$xls.Name; DateRange=$range; Frequency=$usedFreq
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
