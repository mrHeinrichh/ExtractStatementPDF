# Get-Dialog : return the "Statement options" window element, or $null if closed.
function Get-Dialog {
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Statement options")
  return $AE::RootElement.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c)
}
