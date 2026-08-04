# Close-Report : close the Report viewer window if open.
function Close-Report {
  $rep = Get-ReportWindow
  if ($rep) { try { $rep.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch {}; Start-Sleep -Milliseconds 600 }
}
