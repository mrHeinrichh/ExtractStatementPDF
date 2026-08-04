# Export-Csv-FromReport : from the open Report viewer, click the export-format
# dropdown, choose "CSV (comma delimited)", then drive the Save As dialog to save
# to $targetCsv. Returns a status string: saved / overwritten / skipped-existing /
# no-report / no-saveas / save-failed.
#
# The report toolbar isn't automatable, so the two toolbar clicks are fixed
# offsets from the MAXIMIZED window's top-left:
#   export-format dropdown  = window-topleft + (440, 86)
#   "CSV (comma delimited)" = window-topleft + (475, 138)   (2nd item in the menu)
# The Save As dialog itself IS a standard dialog, so its fields are found via UIA.
function Export-Csv-FromReport([string]$targetCsv) {
  $p = Get-Acct
  $rep = Get-ReportWindow; if (-not $rep) { return 'no-report' }
  $wr = New-Object SOA.Win+RECT
  [SOA.Win]::GetWindowRect($p.MainWindowHandle, [ref]$wr) | Out-Null
  $wl = $wr.Left; $wt = $wr.Top
  Click ($wl + 440) ($wt + 86);  Start-Sleep -Milliseconds 800; Shot 'export_menu'
  Click ($wl + 475) ($wt + 138); Start-Sleep -Milliseconds 1200

  $root = $AE::RootElement
  $saCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Save As")
  $sa = $null
  for ($t=0; $t -lt 20; $t++) { $sa = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $saCond); if ($sa) { break }; Start-Sleep -Milliseconds 200 }
  if (-not $sa) { return 'no-saveas' }

  # File name combo (standard dialog AutomationId 1001): clear, type the full path.
  $fnCond = New-Object System.Windows.Automation.PropertyCondition($AE::AutomationIdProperty, "1001")
  $fn = $sa.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $fnCond)
  if ($fn) { $fr = $fn.Current.BoundingRectangle; Click ([int]($fr.X+$fr.Width/2)) ([int]($fr.Y+$fr.Height/2)) }
  Start-Sleep -Milliseconds 250; Send '^a'; Start-Sleep -Milliseconds 80; Send '{DEL}'; Start-Sleep -Milliseconds 80
  Send-Literal $targetCsv; Start-Sleep -Milliseconds 300

  $svCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Save")
  $sv = $sa.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $svCond)
  if ($sv) { $svr = $sv.Current.BoundingRectangle; Click ([int]($svr.X+$svr.Width/2)) ([int]($svr.Y+$svr.Height/2)) } else { Send '{ENTER}' }
  Start-Sleep -Milliseconds 1000

  # "already exists - replace?" prompt: Yes if -Overwrite, else No + Cancel (skip).
  $confirm = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Confirm Save As")))
  if ($confirm) {
    if ($Overwrite) {
      $yes = $confirm.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Yes")))
      if ($yes) { $yr=$yes.Current.BoundingRectangle; Click ([int]($yr.X+$yr.Width/2)) ([int]($yr.Y+$yr.Height/2)) }; Start-Sleep -Milliseconds 800; return 'overwritten'
    } else {
      $no = $confirm.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "No")))
      if ($no) { $nr=$no.Current.BoundingRectangle; Click ([int]($nr.X+$nr.Width/2)) ([int]($nr.Y+$nr.Height/2)) }; Start-Sleep -Milliseconds 500
      $cancel = $sa.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Cancel")))
      if ($cancel) { $cr=$cancel.Current.BoundingRectangle; Click ([int]($cr.X+$cr.Width/2)) ([int]($cr.Y+$cr.Height/2)) }; Start-Sleep -Milliseconds 500
      return 'skipped-existing'
    }
  }
  Start-Sleep -Milliseconds 600
  if (Test-Path -LiteralPath $targetCsv) { return 'saved' }
  return 'save-failed'
}
