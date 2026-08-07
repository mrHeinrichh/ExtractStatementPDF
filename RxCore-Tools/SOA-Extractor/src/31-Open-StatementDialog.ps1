# Open-StatementDialog : make sure the Accounting Print pane is showing, then
# double-click the Statement tile's ICON (~25px above its label) and wait for the
# "Statement options" dialog. Returns the dialog element.
#
# Robust across machines/monitors:
#   * clicks the LEFTMOST "Accounting Print" sidebar item relative to THIS window
#     (not a hardcoded screen-left assumption), and retries;
#   * if the pane still can't be found, throws a diagnostic listing what it DID see
#     (window rect, left-column labels, custom panes) so a failing PC is easy to debug.
function Open-StatementDialog {
  $p = Focus-Acct
  $root = $AE::FromHandle($p.MainWindowHandle)
  $wr = New-Object SOA.Win+RECT
  [SOA.Win]::GetWindowRect($p.MainWindowHandle, [ref]$wr) | Out-Null

  # Click the "Accounting Print" sidebar item (leftmost one within this window),
  # then wait for the PrintAccountingReports pane. Retry a few times.
  $par = $null
  for ($attempt = 0; $attempt -lt 4 -and -not $par; $attempt++) {
    $apCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Accounting Print")
    $aps = @($root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $apCond) |
             Where-Object { $_.Current.ClassName -eq 'TextBlock' -and $_.Current.BoundingRectangle.Width -gt 0 } |
             Sort-Object { $_.Current.BoundingRectangle.X })
    if ($aps.Count -gt 0) {
      $b = $aps[0].Current.BoundingRectangle
      Click ([int]($b.X + $b.Width/2)) ([int]($b.Y + $b.Height/2))
    }
    for ($t = 0; $t -lt 10; $t++) {
      $parCond = New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, "PrintAccountingReports")
      $par = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $parCond)
      if ($par) { break }
      Start-Sleep -Milliseconds 300
    }
  }

  if (-not $par) {
    # Build a diagnostic of what this window actually exposes.
    $labels = @($root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                 (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text))) |
               Where-Object { $_.Current.BoundingRectangle.X -lt ($wr.Left + 320) -and $_.Current.Name } |
               ForEach-Object { $_.Current.Name } | Select-Object -Unique -First 25)
    $customs = @($root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                 (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Custom))) |
               ForEach-Object { $_.Current.ClassName } | Where-Object { $_ } | Select-Object -Unique)
    $elevNote = if (Test-ElevationMismatch) {
      "`n  *** ELEVATION MISMATCH: Accounting is elevated but this tool is not - run this tool as administrator, or restart Accounting without admin. ***"
    } else {
      "  (this tool elevated=$(Test-SelfElevated); Accounting readable=$(Test-AcctAccessible))"
    }
    throw ("Accounting Print pane not found.`n" +
           "  Window rect: $($wr.Left),$($wr.Top)-$($wr.Right),$($wr.Bottom)`n" +
           "  Left-column labels seen: " + ($labels -join ', ') + "`n" +
           "  Custom panes seen: " + ($customs -join ', ') + "`n" +
           $elevNote + "`n" +
           "  -> Make sure the chosen Accounting window is logged in and on the General > Accounting Print tab.")
  }

  $stCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Statement")
  $st = $par.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $stCond)
  if (-not $st) { throw "'Statement' tile not found inside the Accounting Print pane." }
  $r = $st.Current.BoundingRectangle; $cx = [int]($r.X + $r.Width/2); $iy = [int]($r.Y - 25)
  for ($i = 0; $i -lt 4; $i++) {
    Click $cx $iy -Double
    for ($t = 0; $t -lt 15; $t++) { $d = Get-Dialog; if ($d) { return $d }; Start-Sleep -Milliseconds 200 }
    Start-Sleep -Milliseconds 400
  }
  throw "Statement options dialog did not open."
}
