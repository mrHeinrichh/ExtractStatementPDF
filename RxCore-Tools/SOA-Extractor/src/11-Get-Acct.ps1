# Get-Acct : return the Accounting process to drive. Prefers the one chosen via
# Choose-Acct ($script:AcctPid); otherwise the first running "Accounting" process.
function Get-Acct {
  if ($script:AcctPid) {
    $p = Get-Process -Id $script:AcctPid -ErrorAction SilentlyContinue
    if ($p) { return $p }
  }
  $p = Get-Process Accounting -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $p) { throw "Accounting.exe is not running. Open it and log in first." }
  return $p
}
