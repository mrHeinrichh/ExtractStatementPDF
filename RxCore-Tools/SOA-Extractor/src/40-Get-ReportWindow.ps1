# Get-ReportWindow : return the "Report" viewer window element, or $null.
function Get-ReportWindow {
  $root = $AE::FromHandle((Get-Acct).MainWindowHandle)
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Report")
  return $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c)
}
