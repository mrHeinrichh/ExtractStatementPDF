# Assert-AcctLoggedIn : verify Accounting.exe is running AND logged in.
#
# When logged in, one of the app's top-level windows is titled
# "Accounting( Version ... ) Company: <co> User: <name>". At the login screen no
# window contains "User:". We scan ALL of the process's top-level windows rather
# than just MainWindowTitle, because MainWindowTitle can read empty transiently
# (e.g. while a child window/dialog has focus) and give a false "not logged in".
#
# Shows a message box and returns $false if not ready; otherwise $true.

function Get-AcctWindowTitles {
  $procs = Get-Process Accounting -ErrorAction SilentlyContinue
  if (-not $procs) { return @() }
  $titles = New-Object System.Collections.Generic.List[string]
  $desktop = $AE::RootElement
  foreach ($proc in $procs) {
    try {
      $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, [int]$proc.Id)
      $wins = $desktop.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)
      foreach ($w in $wins) { $titles.Add([string]$w.Current.Name) }
    } catch {}
    if ($proc.MainWindowTitle) { $titles.Add([string]$proc.MainWindowTitle) }
  }
  return $titles
}

function Assert-AcctLoggedIn {
  $procs = Get-Process Accounting -ErrorAction SilentlyContinue
  if (-not $procs) {
    Show-Message "Accounting is not open.`n`nPlease open Accounting.exe and log in (user hfabros), then run this again."
    return $false
  }
  $titles = Get-AcctWindowTitles
  if (@($titles | Where-Object { $_ -match 'User:' }).Count -gt 0) { return $true }
  Show-Message "You are not logged in to Accounting.`n`nPlease log in on the Accounting window (user hfabros), then run this again."
  return $false
}
