# Focus-Acct : bring Accounting to the front and MAXIMIZE it, so all screen
# coordinates used later (toolbar/export offsets) are deterministic.
function Focus-Acct {
  $p = Get-Acct
  [SOA.Win]::ShowWindow($p.MainWindowHandle, 3)       | Out-Null   # SW_MAXIMIZE
  [SOA.Win]::SetForegroundWindow($p.MainWindowHandle) | Out-Null
  [SOA.Win]::BringWindowToTop($p.MainWindowHandle)    | Out-Null
  Start-Sleep -Milliseconds 400
  return $p
}
