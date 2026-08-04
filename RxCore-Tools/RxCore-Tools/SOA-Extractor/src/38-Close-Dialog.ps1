# Close-Dialog : close the Statement dialog (via the window Close pattern, or
# the "No" button as a fallback).
function Close-Dialog {
  $d = Get-Dialog
  if ($d) { try { $d.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch { $xy=Field-XY 'No'; Click $xy[0] $xy[1] }; Start-Sleep -Milliseconds 500 }
}
