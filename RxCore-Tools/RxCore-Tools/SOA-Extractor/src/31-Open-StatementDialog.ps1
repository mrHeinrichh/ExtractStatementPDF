# Open-StatementDialog : go to the Accounting Print tab and double-click the
# Statement tile's ICON (~25px above its label - the label itself isn't clickable),
# then wait for the "Statement options" dialog. Returns the dialog element.
function Open-StatementDialog {
  $p = Focus-Acct; $root = $AE::FromHandle($p.MainWindowHandle)
  $apCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Accounting Print")
  foreach ($ap in $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $apCond)) {
    if ($ap.Current.ClassName -eq 'TextBlock' -and $ap.Current.BoundingRectangle.X -lt 200) {
      $b = $ap.Current.BoundingRectangle; Click ([int]($b.X+$b.Width/2)) ([int]($b.Y+$b.Height/2)); Start-Sleep -Milliseconds 600; break
    }
  }
  $parCond = New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, "PrintAccountingReports")
  $par = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $parCond)
  if (-not $par) { throw "Accounting Print pane not found." }
  $stCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Statement")
  $st = $par.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $stCond)
  if (-not $st) { throw "'Statement' tile not found." }
  $r = $st.Current.BoundingRectangle; $cx = [int]($r.X+$r.Width/2); $iy = [int]($r.Y-25)
  for ($i=0; $i -lt 4; $i++) {
    Click $cx $iy -Double
    for ($t=0; $t -lt 15; $t++) { $d = Get-Dialog; if ($d) { return $d }; Start-Sleep -Milliseconds 200 }
    Start-Sleep -Milliseconds 400
  }
  throw "Statement options dialog did not open."
}
