# Reset-State : close any leftover Report window and Statement dialog so the
# next customer starts from a clean screen.
function Reset-State {
  for ($k = 0; $k -lt 3; $k++) {
    $rep = Get-ReportWindow
    if (-not $rep) { break }
    try { $rep.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch {}
    Start-Sleep -Milliseconds 600
  }
  $d = Get-Dialog
  if ($d) {
    try { $d.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch {}
    Start-Sleep -Milliseconds 400
  }
}
