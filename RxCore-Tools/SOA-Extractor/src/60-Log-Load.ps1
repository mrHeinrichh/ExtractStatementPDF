# Load-Log : read the persistent run log (_SOA_Log.csv) into an ordered hashtable
# keyed by File name. Each value is the last-known row for that .xls. Empty if the
# log doesn't exist yet (first run). This is what makes the tool resumable: files
# already marked DONE here are skipped on the next run.
function Load-Log([string]$logCsv) {
  $log = [ordered]@{}
  if (-not (Test-Path -LiteralPath $logCsv)) { return $log }
  try {
    foreach ($row in (Import-Csv -LiteralPath $logCsv -Encoding UTF8)) {
      if ($row.File) { $log[$row.File] = $row }
    }
  } catch {}
  return $log
}
