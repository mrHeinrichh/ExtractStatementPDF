# =====================  MAIN PROGRAM  =====================
# Test hooks (optional): CM_XLS (folder), CM_REF (customer csv), CM_OUT (write target).

# 1) folder of .xls files
if ($env:CM_XLS) { $xlsFolder = $env:CM_XLS } else { $xlsFolder = Pick-Folder }
if (-not $xlsFolder) { Write-Host "No folder selected. Exiting."; return }
if (-not (Test-Path -LiteralPath $xlsFolder)) { Write-Host "Folder not found: $xlsFolder"; return }

# 2) the Customer file exported from Contact (browse; opens in Documents)
if ($env:CM_REF) { $refPath = $env:CM_REF } else { $refPath = Browse-File "Select the Customer file exported from Contact" }
if (-not $refPath) { Write-Host "No Customer file selected. Exiting."; return }

# 3) write target = the SAME file (in place). CM_OUT can redirect it (used for tests).
$outPath = if ($env:CM_OUT) { $env:CM_OUT } else { $refPath }

Write-Host "xls folder    : $xlsFolder"
Write-Host "Customer file : $refPath"
Write-Host "Writing to    : $outPath  (in place)`n"

$parsed = Read-CustomerExport $refPath
Write-Host ("Customer rows : {0}   (Name col {1}{2})" -f $parsed.Data.Count, $parsed.INameIdx, $(if ($parsed.IMatchIdx -ge 0) { '; updating existing MatchedXlsName column' } else { '' }))

$xlsNames = Read-XlsNames $xlsFolder
Write-Host ("Extracted .xls names : {0}" -f $xlsNames.Count)

$res = Match-Names $parsed $xlsNames
Write-Host ("Matched rows  : {0}   Unmatched .xls names : {1}" -f $res.Matched.Count, $res.Unmatched.Count)

try {
  Write-Annotated $parsed $res.Matched $outPath
  Write-Host "`nDone. Matched names written into: $outPath" -ForegroundColor Green
} catch {
  Write-Host "`n$($_.Exception.Message)" -ForegroundColor Red
  Show-Message $_.Exception.Message
  return
}

if ($res.Unmatched.Count) {
  Write-Host "`nUnmatched .xls names (no Contact row found):" -ForegroundColor Yellow
  $res.Unmatched | ForEach-Object { Write-Host "  $_" }
}
