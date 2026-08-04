# Get-Acct : return the running Accounting.exe process, or throw if not running.
function Get-Acct {
  $p = Get-Process Accounting -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $p) { throw "Accounting.exe is not running. Open it and log in first." }
  return $p
}
